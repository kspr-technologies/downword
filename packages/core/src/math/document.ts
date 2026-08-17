/**
 * The math pass: walk a parsed document, fill every equation's `omml` slot,
 * hand the renderer a tree it can render synchronously.
 *
 * ```ts
 * const doc = parseMarkdown(md, { plugins: [mathMarkdownIt()] });
 * const { document } = await convertDocumentMath(doc);
 * const file = renderDocument(document);
 * ```
 *
 * It is the same seam `resolveDocumentImages` occupies — the async part of the
 * pipeline, between a pure parse and a pure render — with one difference that
 * matters. Images are handed to the renderer in a `Map` keyed by node, because
 * `ImageNode.resolved` is a slot the *renderer* reads through the map. An
 * equation's OMML has no such map: `renderDocument` reads `node.omml` directly.
 * So this pass returns a **new document** with the slots filled.
 *
 * ## The rewrite preserves identity everywhere it can
 *
 * Only the nodes on the path from the root to an equation are rebuilt; every
 * other node in the tree is the *same object*. That is not an optimisation, it
 * is a correctness requirement: `prepareHighlights` and `resolveDocumentImages`
 * key their maps by node identity, so cloning a code block or a picture that
 * has nothing to do with maths would silently throw its resolution away. A
 * document with no equations comes back as the very same object.
 *
 * ## Degradation ladder
 *
 * Every failure moves one rung down, never off the end:
 *
 * ```text
 *   picture ──▶ native equation (OMML) ──▶ the TeX source as literal text
 * ```
 *
 * The bottom rung is the renderer's existing behaviour for an unconverted
 * equation, so the document is always produced and the reader always sees
 * something. Nothing in this file throws.
 */

import type {
  BlockNode,
  DocumentNode,
  InlineNode,
  ListItemNode,
  MathBlockNode,
  MathInlineNode,
  TableCellNode,
  TableRowNode,
} from "../model.js";
import { assertNever, nodesOfType, paragraph } from "../model.js";
import { createMathConverter } from "./convert.js";
import { mathImageNode, rasterizeMath } from "./image.js";
import { resolveMathOptions, type MathOptions } from "./options.js";
import { createMathWarningSink, type MathWarning } from "./warnings.js";

/** How many equations to rasterise at once. Matches the image pass. */
const RASTER_CONCURRENCY = 4;

/** What {@link convertDocumentMath} produced. */
export interface MathDocumentConversion {
  /** The document to render. The same object when nothing changed. */
  readonly document: DocumentNode;
  /** Every problem, in the order it was found. */
  readonly diagnostics: readonly MathWarning[];
  /** Equations that became a native Word equation. */
  readonly ommlCount: number;
  /** Equations that became a picture. */
  readonly imageCount: number;
  /** Equations left as literal TeX. */
  readonly unconvertedCount: number;
}

/** What one equation turns into. */
type InlineReplacement = ReadonlyMap<MathInlineNode, InlineNode>;
type BlockReplacement = ReadonlyMap<MathBlockNode, BlockNode>;

/**
 * Converts every `$…$` and `$$…$$` in a document.
 *
 * @param document - A parsed document. Not modified.
 * @param options - See {@link MathOptions}. Every field has a default.
 * @returns The converted document plus what happened. Never rejects.
 * @throws {import("../errors.js").DownwordError} `"invalid-options"`, and only
 *   for an option value it cannot use — never for anything in the document.
 */
export async function convertDocumentMath(
  document: DocumentNode,
  options: MathOptions = {},
): Promise<MathDocumentConversion> {
  const resolved = resolveMathOptions(options);

  const diagnostics: MathWarning[] = [];
  const handler = resolved.onWarning;
  const sink = createMathWarningSink((warning) => {
    diagnostics.push(warning);
    handler?.(warning);
  });

  const inlineNodes = nodesOfType(document, "mathInline");
  const blockNodes = nodesOfType(document, "mathBlock");

  const nothing: MathDocumentConversion = {
    document,
    diagnostics,
    ommlCount: 0,
    imageCount: 0,
    unconvertedCount: 0,
  };

  // "off" is a real answer, not a degenerate one: the renderer already puts the
  // TeX source in the document, and skipping the pass means neither engine is
  // ever imported.
  if (resolved.math === "off") return nothing;
  if (inlineNodes.length === 0 && blockNodes.length === 0) return nothing;

  const converter = await createMathConverter(resolved, sink);
  if (converter === null) {
    return { ...nothing, unconvertedCount: inlineNodes.length + blockNodes.length };
  }

  if (resolved.math === "image" && resolved.rasterizer === null) {
    sink.report(
      "image-no-rasterizer",
      "",
      false,
      `math: "image" needs a rasterizer to turn each equation into the raster twin OOXML ` +
        `requires, and none was supplied — drawing MathML needs a browser, so this mode is ` +
        `browser-only. Every equation was written as a native Word equation instead`,
    );
  }

  const inline = new Map<MathInlineNode, InlineNode>();
  const block = new Map<MathBlockNode, BlockNode>();
  let ommlCount = 0;
  let imageCount = 0;
  let unconvertedCount = 0;

  /** One equation, whichever kind of node it came from. */
  interface Task {
    readonly tex: string;
    readonly display: boolean;
    readonly apply: (replacement: { omml: string | null; picture: InlineNode | null }) => void;
  }

  const tasks: Task[] = [
    ...inlineNodes.map((node): Task => ({
      tex: node.value,
      display: false,
      apply: ({ omml, picture }) => {
        if (picture !== null) inline.set(node, picture);
        else if (omml !== null) inline.set(node, { ...node, omml });
      },
    })),
    ...blockNodes.map((node): Task => ({
      tex: node.value,
      display: true,
      apply: ({ omml, picture }) => {
        if (picture !== null) block.set(node, paragraph([picture]));
        else if (omml !== null) block.set(node, { ...node, omml });
      },
    })),
  ];

  const runTask = async (task: Task): Promise<void> => {
    const { mathml, omml } = converter.convert(task.tex, task.display);

    if (resolved.math === "image" && resolved.rasterizer !== null && mathml !== null) {
      const picture = await rasterizeMath(mathml, task.tex, task.display, resolved, sink);
      if (picture !== null) {
        imageCount += 1;
        task.apply({ omml: null, picture: mathImageNode(task.tex, picture) });
        return;
      }
    }

    if (omml !== null) ommlCount += 1;
    else unconvertedCount += 1;
    task.apply({ omml, picture: null });
  };

  // Only the rasteriser is actually asynchronous; in "omml" mode this settles
  // on the first microtask.
  let next = 0;
  const worker = async (): Promise<void> => {
    for (;;) {
      const index = next;
      next += 1;
      const task = tasks[index];
      if (task === undefined) return;
      await runTask(task);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(RASTER_CONCURRENCY, Math.max(1, tasks.length)) }, () => worker()),
  );

  return {
    document: rewriteDocument(document, inline, block),
    diagnostics,
    ommlCount,
    imageCount,
    unconvertedCount,
  };
}

/* -------------------------------------------------------------------------- */
/* The rewrite                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Maps `items` and returns the *original array* when every element came back
 * identical, so an untouched subtree keeps its object identity.
 */
function mapPreservingIdentity<T>(items: readonly T[], map: (item: T) => T): readonly T[] {
  let out: T[] | null = null;
  for (const [index, item] of items.entries()) {
    const next = map(item);
    if (next !== item && out === null) out = items.slice(0, index);
    if (out !== null) out.push(next);
  }
  return out ?? items;
}

function rewriteInline(node: InlineNode, inline: InlineReplacement): InlineNode {
  return node.type === "mathInline" ? (inline.get(node) ?? node) : node;
}

function rewriteCell(cell: TableCellNode, inline: InlineReplacement): TableCellNode {
  const children = mapPreservingIdentity(cell.children, (child) => rewriteInline(child, inline));
  return children === cell.children ? cell : { ...cell, children };
}

function rewriteRow(row: TableRowNode, inline: InlineReplacement): TableRowNode {
  const children = mapPreservingIdentity(row.children, (cell) => rewriteCell(cell, inline));
  return children === row.children ? row : { ...row, children };
}

function rewriteItem(
  item: ListItemNode,
  inline: InlineReplacement,
  block: BlockReplacement,
): ListItemNode {
  const children = rewriteBlocks(item.children, inline, block);
  return children === item.children ? item : { ...item, children };
}

function rewriteBlocks(
  blocks: readonly BlockNode[],
  inline: InlineReplacement,
  block: BlockReplacement,
): readonly BlockNode[] {
  return mapPreservingIdentity(blocks, (child) => rewriteBlock(child, inline, block));
}

/**
 * Rebuilds one block, or returns it untouched.
 *
 * An exhaustive switch on purpose: a node kind added to the model must state
 * whether equations can hide inside it, rather than silently losing them.
 */
function rewriteBlock(
  node: BlockNode,
  inline: InlineReplacement,
  block: BlockReplacement,
): BlockNode {
  switch (node.type) {
    case "mathBlock":
      return block.get(node) ?? node;

    case "paragraph":
    case "heading": {
      const children = mapPreservingIdentity(node.children, (child) =>
        rewriteInline(child, inline),
      );
      return children === node.children ? node : { ...node, children };
    }

    case "list": {
      const children = mapPreservingIdentity(node.children, (item) =>
        rewriteItem(item, inline, block),
      );
      return children === node.children ? node : { ...node, children };
    }

    case "table": {
      const children = mapPreservingIdentity(node.children, (row) => rewriteRow(row, inline));
      return children === node.children ? node : { ...node, children };
    }

    case "blockquote":
    case "footnoteDefinition": {
      const children = rewriteBlocks(node.children, inline, block);
      return children === node.children ? node : { ...node, children };
    }

    case "codeBlock":
    case "thematicBreak":
    case "htmlBlock":
      return node;

    default:
      return assertNever(node, "block node");
  }
}

function rewriteDocument(
  document: DocumentNode,
  inline: InlineReplacement,
  block: BlockReplacement,
): DocumentNode {
  if (inline.size === 0 && block.size === 0) return document;
  const children = rewriteBlocks(document.children, inline, block);
  return children === document.children ? document : { ...document, children };
}
