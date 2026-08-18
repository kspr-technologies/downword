/**
 * The pass: ```` ```mermaid ```` fences in, embedded pictures out.
 *
 * ### Shape of the thing
 *
 * This is a **document transform**, not a renderer change. It takes a parsed
 * `DocumentNode` and returns a new one in which every diagram fence has become
 * an image paragraph (plus, when the fence carried one, a caption paragraph),
 * with the bytes already in the node's `resolved` slot. Nothing downstream
 * needs to know mermaid exists: the renderer sees an ordinary image, sizes it
 * to the text column, centres it in a `Figure` paragraph and writes the alt
 * text into `wp:docPr` exactly as it would for `![](diagram.png)`.
 *
 * That is also why it is additive: the model union is untouched, no node type
 * was added, and `src/render/**` is not involved.
 *
 * ### Identity is preserved
 *
 * Untouched subtrees come back as the *same objects*, and untouched arrays as
 * the same arrays. `prepareHighlights` and `prepareImages` key their maps by
 * node identity, so a pass that rebuilt the whole tree would quietly invalidate
 * a map the caller had already computed. Run this first anyway — it is the
 * cheapest order and the one the docs show — but a caller who does not is not
 * punished for it.
 *
 * ### Serial, not parallel
 *
 * Diagrams are rendered one at a time. mermaid keeps module-level state, mounts
 * its measuring element in the shared document, and is explicitly not
 * re-entrant; `prepareHighlights` runs serially for the same reason.
 */

import { isValidResolvedImage } from "../images/decode.js";
import { DEFAULT_CONTENT_WIDTH_TWIPS, fitToWidth } from "../images/layout.js";
import { probeSvg } from "../images/svg.js";
import type { ImageRasterizer } from "../images/types.js";
import {
  image,
  italic,
  paragraph,
  text,
  type BlockNode,
  type CodeBlockNode,
  type DocumentNode,
  type ListItemNode,
  type ListNode,
  type ResolvedImage,
} from "../model.js";
import { invalid } from "../validate.js";
import { createBrowserMermaidRenderer, isEngineUnavailable, type MermaidConfig } from "./engine.js";
import { detectDomSupport } from "./environment.js";
import { isMermaidFence, readMermaidFence } from "./fence.js";
import { createCanvasRasterizer } from "./raster.js";
import {
  mermaidWarningSeverity,
  type MermaidDiagnosticCode,
  type MermaidRenderer,
  type MermaidWarning,
  type MermaidWarningHandler,
} from "./types.js";

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

/** Options for {@link renderMermaid}. */
export interface MermaidOptions {
  /**
   * Turns diagram source into SVG. Defaults to
   * {@link createBrowserMermaidRenderer}, which dynamically imports mermaid —
   * so the default needs a DOM and the optional peer dependency.
   *
   * Supply your own to use a mermaid you already have, or to render somewhere
   * this package cannot reach.
   */
  readonly renderer?: MermaidRenderer | undefined;
  /**
   * Turns that SVG into the raster OOXML requires. Defaults to
   * {@link createCanvasRasterizer}, which also needs a DOM.
   *
   * In Node, `resvg`/`sharp`/`@napi-rs/canvas` wrapped in an
   * {@link ImageRasterizer} is what makes the rest of this pass work headlessly
   * — but only in combination with a `renderer`, because mermaid itself cannot
   * run without a document.
   */
  readonly rasterizer?: ImageRasterizer | undefined;
  /**
   * Raster pixels per display pixel. Defaults to `2`.
   *
   * Only reaches the default rasteriser; a supplied one decides for itself.
   */
  readonly scale?: number | undefined;
  /**
   * Usable text width in twips, used to decide how large the raster needs to
   * be. Defaults to {@link DEFAULT_CONTENT_WIDTH_TWIPS} (A4, one-inch margins).
   *
   * It is a *target*, not a clamp: the renderer fits the picture to the real
   * page at layout time, when the real geometry is known.
   */
  readonly maxWidthTwips?: number | undefined;
  /**
   * What lands in the document. Defaults to `"png"`.
   *
   * - `"png"` — the raster alone. Every reader draws it, the file is smaller,
   *   and at 2x it is sharp on screen and in print.
   * - `"svg"` — the SVG *and* the raster twin OOXML demands. Word 2016+ draws
   *   the vector, so it stays crisp at any zoom; everything else falls back to
   *   the same PNG it would have got. Costs both copies.
   */
  readonly embed?: "png" | "svg" | undefined;
  /**
   * Emit the fence's caption as a paragraph under the picture. Defaults to
   * `true`.
   *
   * The caption comes from the info string (```` ```mermaid My caption ````) or
   * from mermaid's own `title:` frontmatter, and is rendered as an italic
   * paragraph — downword has no `Caption` style to point at, and inventing one
   * from a plugin would put a style in `styles.xml` that the theme does not
   * know about. With this off the text still becomes the picture's alt text.
   */
  readonly caption?: boolean | undefined;
  /** mermaid configuration for the default renderer. See {@link createBrowserMermaidRenderer}. */
  readonly config?: MermaidConfig | undefined;
  /**
   * Give up on a diagram after this many milliseconds. Defaults to `10000`.
   *
   * mermaid parses untrusted text and can take pathologically long on
   * adversarial input; a conversion that never finishes is worse than a
   * conversion with one fence left in it. `0` disables the timeout.
   */
  readonly timeoutMs?: number | undefined;
  /** Called once per problem. Also collected in {@link MermaidResult.warnings}. */
  readonly onWarning?: MermaidWarningHandler | undefined;
}

/** What {@link renderMermaid} produces. */
export interface MermaidResult {
  /**
   * The rewritten document — or the *same object*, when nothing changed.
   *
   * Always usable: on every failure path this is a document that still contains
   * every diagram's source as a fenced code block.
   */
  readonly document: DocumentNode;
  /** Diagram fences found. */
  readonly diagrams: number;
  /** Fences that became pictures. */
  readonly rendered: number;
  /** Every problem, in document order. See {@link MermaidWarning}. */
  readonly warnings: readonly MermaidWarning[];
}

const DEFAULT_TIMEOUT_MS = 10_000;

/* -------------------------------------------------------------------------- */
/* Tree rewriting                                                              */
/* -------------------------------------------------------------------------- */

/** Diagram fence -> the blocks that replace it. */
type Replacements = ReadonlyMap<CodeBlockNode, readonly BlockNode[]>;

/**
 * Rewrites a list of blocks, recursing into the four containers a fence can sit
 * in.
 *
 * Returns the input array itself when nothing below it changed, which is what
 * keeps node identity stable for the parts of the document this pass did not
 * touch. Table cells hold inline content only, so they cannot contain a fence
 * and are not walked.
 */
function rewriteBlocks(
  blocks: readonly BlockNode[],
  replacements: Replacements,
): readonly BlockNode[] {
  let changed = false;
  const out: BlockNode[] = [];

  for (const block of blocks) {
    const replacement = block.type === "codeBlock" ? replacements.get(block) : undefined;
    if (replacement !== undefined) {
      out.push(...replacement);
      changed = true;
      continue;
    }

    const rewritten = rewriteBlock(block, replacements);
    if (rewritten !== block) changed = true;
    out.push(rewritten);
  }

  return changed ? out : blocks;
}

function rewriteListItems(
  items: readonly ListItemNode[],
  replacements: Replacements,
): readonly ListItemNode[] {
  let changed = false;
  const out: ListItemNode[] = [];

  for (const item of items) {
    const children = rewriteBlocks(item.children, replacements);
    if (children === item.children) {
      out.push(item);
      continue;
    }
    changed = true;
    out.push({ ...item, children });
  }

  return changed ? out : items;
}

function rewriteBlock(block: BlockNode, replacements: Replacements): BlockNode {
  switch (block.type) {
    case "blockquote":
    case "footnoteDefinition": {
      const children = rewriteBlocks(block.children, replacements);
      return children === block.children ? block : { ...block, children };
    }
    case "list": {
      const children = rewriteListItems(block.children, replacements);
      return children === block.children ? block : ({ ...block, children } satisfies ListNode);
    }
    default:
      // Paragraphs, headings, tables, code, rules and HTML hold no block
      // children a fence could hide in.
      return block;
  }
}

/** Every diagram fence in the document, in document order. */
function findFences(doc: DocumentNode): readonly CodeBlockNode[] {
  const found: CodeBlockNode[] = [];

  const walkBlocks = (blocks: readonly BlockNode[]): void => {
    for (const block of blocks) {
      if (isMermaidFence(block)) {
        found.push(block);
        continue;
      }
      switch (block.type) {
        case "blockquote":
        case "footnoteDefinition":
          walkBlocks(block.children);
          break;
        case "list":
          for (const item of block.children) walkBlocks(item.children);
          break;
        default:
          break;
      }
    }
  };

  walkBlocks(doc.children);
  return found;
}

/**
 * How many diagram fences a document contains.
 *
 * Cheap, synchronous and DOM-free, so a host can decide whether to offer
 * diagram rendering at all — or a CLI can say "3 diagrams were left as code
 * blocks" without loading anything.
 */
export function countMermaidDiagrams(doc: DocumentNode): number {
  return findFences(doc).length;
}

/* -------------------------------------------------------------------------- */
/* The pass                                                                    */
/* -------------------------------------------------------------------------- */

function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Rejects after `ms`, so a renderer that never settles cannot hang a conversion. */
async function withTimeout<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  if (ms <= 0) return work;
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      work,
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          reject(new Error(`${what} took longer than ${ms}ms`));
        }, ms);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

function positive(value: number | undefined, option: string, fallback: number): number {
  if (value === undefined) return fallback;
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw invalid(`options.${option} must be a finite number > 0, got ${String(value)}`);
  }
  return value;
}

function nonNegative(value: number | undefined, option: string, fallback: number): number {
  if (value === undefined) return fallback;
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw invalid(`options.${option} must be a finite number >= 0, got ${String(value)}`);
  }
  return value;
}

/**
 * Turns every ```` ```mermaid ```` fence in a document into an embedded picture.
 *
 * ```ts
 * import { parseMarkdown, renderDocument } from "@ksprtech/downword";
 * import { renderMermaid } from "@ksprtech/downword/plugins/mermaid";
 *
 * const parsed = parseMarkdown(markdown);
 * const { document, warnings } = await renderMermaid(parsed);
 * const file = renderDocument(document);
 * ```
 *
 * **Never rejects, and never loses a diagram.** Anything that goes wrong —
 * no DOM, no mermaid installed, a diagram that does not parse, a canvas that
 * refuses — leaves that fence exactly as it was, so the document still contains
 * the source, and reports why through {@link MermaidOptions.onWarning} and
 * {@link MermaidResult.warnings}.
 *
 * @param doc - A parsed document. Not modified.
 * @param options - See {@link MermaidOptions}. Every field has a default.
 * @returns The rewritten document, counts and diagnostics.
 * @throws {import("../errors.js").DownwordError} with code `"invalid-options"`
 *   for an out-of-range `scale`, `maxWidthTwips` or `timeoutMs`, and nothing else.
 */
export async function renderMermaid(
  doc: DocumentNode,
  options: MermaidOptions = {},
): Promise<MermaidResult> {
  const warnings: MermaidWarning[] = [];
  const handler = options.onWarning;
  const warn = (code: MermaidDiagnosticCode, diagram: number | null, message: string): void => {
    const warning: MermaidWarning = {
      code,
      severity: mermaidWarningSeverity(code),
      diagram,
      message,
    };
    warnings.push(warning);
    try {
      handler?.(warning);
    } catch {
      // A host whose logger throws must not take the conversion down with it.
    }
  };

  const scale = positive(options.scale, "scale", 2);
  const maxWidthTwips = positive(
    options.maxWidthTwips,
    "maxWidthTwips",
    DEFAULT_CONTENT_WIDTH_TWIPS,
  );
  const timeoutMs = nonNegative(options.timeoutMs, "timeoutMs", DEFAULT_TIMEOUT_MS);
  const embed = options.embed ?? "png";
  const wantsCaption = options.caption ?? true;

  const fences = findFences(doc);
  if (fences.length === 0) {
    return { document: doc, diagrams: 0, rendered: 0, warnings };
  }

  const support = detectDomSupport();
  const renderer =
    options.renderer ??
    (support.renderer
      ? createBrowserMermaidRenderer({
          ...(options.config === undefined ? {} : { config: options.config }),
        })
      : null);
  const rasterizer =
    options.rasterizer ?? (support.rasterizer ? createCanvasRasterizer({ scale }) : null);

  const count = `${fences.length} mermaid ${fences.length === 1 ? "diagram" : "diagrams"}`;
  const kept = `left as ${fences.length === 1 ? "a fenced code block" : "fenced code blocks"}`;

  if (renderer === null) {
    warn(
      "no-dom",
      null,
      `${count} ${kept}: ${support.reason}. Convert in a browser, or pass renderMermaid() ` +
        `your own \`renderer\` (and \`rasterizer\`) — for example mermaid driven through jsdom, ` +
        `or a headless browser.`,
    );
    return { document: doc, diagrams: fences.length, rendered: 0, warnings };
  }

  if (rasterizer === null) {
    warn(
      "rasterizer-unavailable",
      null,
      `${count} ${kept}: there is no rasterizer, and the default canvas one cannot be built here` +
        `${support.reason === "" ? "" : ` (${support.reason})`}. OOXML cannot hold a bare SVG — ` +
        `it stores a raster twin alongside it — so pass renderMermaid() a \`rasterizer\`: a ` +
        `canvas in the browser, resvg or sharp in Node.`,
    );
    return { document: doc, diagrams: fences.length, rendered: 0, warnings };
  }

  const replacements = new Map<CodeBlockNode, readonly BlockNode[]>();
  let rendered = 0;

  for (const [index, fence] of fences.entries()) {
    const diagram = index + 1;
    const info = readMermaidFence(fence);

    if (info.source === "") {
      warn("empty-diagram", diagram, `mermaid diagram ${diagram} is empty; the fence was kept`);
      continue;
    }

    let svg: string;
    let declared: { width: number | null; height: number | null };
    try {
      const output = await withTimeout(
        renderer.render({
          source: info.source,
          // Deterministic, so two conversions of one document produce the same
          // SVG bytes: mermaid writes this id into the markup it returns.
          id: `downword-mermaid-${diagram}`,
          diagram,
        }),
        timeoutMs,
        `rendering mermaid diagram ${diagram}`,
      );
      svg = output.svg;
      declared = { width: output.width ?? null, height: output.height ?? null };
    } catch (error: unknown) {
      if (isEngineUnavailable(error)) {
        // One fact about the whole run, not one broken diagram: report it once
        // and stop, rather than failing identically N times.
        warn("engine-unavailable", null, `${count} ${kept}: ${describe(error)}`);
        break;
      }
      warn(
        "render-failed",
        diagram,
        `mermaid could not draw diagram ${diagram} (${info.kind}); the fence was kept: ${describe(error)}`,
      );
      continue;
    }

    const bytes = new TextEncoder().encode(svg);
    const probe = probeSvg(bytes);
    if (probe === null) {
      warn(
        "invalid-svg",
        diagram,
        `mermaid returned something that is not an SVG for diagram ${diagram}; the fence was kept`,
      );
      continue;
    }

    const intrinsic = {
      width: declared.width ?? probe.width,
      height: declared.height ?? probe.height,
    };
    // The size the picture will occupy. The raster carries `scale` times as
    // many pixels; OOXML records the two separately.
    const display = fitToWidth(intrinsic, maxWidthTwips);

    let raster;
    try {
      raster = await withTimeout(
        rasterizer.rasterize({
          data: bytes,
          format: "svg",
          src: `mermaid:diagram-${diagram}`,
          intrinsicWidth: probe.intrinsic ? intrinsic.width : null,
          intrinsicHeight: probe.intrinsic ? intrinsic.height : null,
          targetWidth: display.width,
          targetHeight: display.height,
        }),
        timeoutMs,
        `rasterising mermaid diagram ${diagram}`,
      );
    } catch (error: unknown) {
      warn(
        "rasterize-failed",
        diagram,
        `the rasteriser threw on diagram ${diagram}; the fence was kept: ${describe(error)}`,
      );
      continue;
    }

    if (raster === null) {
      warn(
        "rasterize-failed",
        diagram,
        `the rasteriser produced no bytes for diagram ${diagram}; the fence was kept`,
      );
      continue;
    }

    const resolved: ResolvedImage =
      embed === "svg"
        ? {
            format: "svg",
            // Bytes, never markup: ImageRun runs any string through atob().
            data: bytes,
            width: display.width,
            height: display.height,
            fallback: raster,
          }
        : {
            format: raster.format,
            data: raster.data,
            // The *display* size, not the raster's pixel count: the bytes are
            // `scale` times denser on purpose.
            width: display.width,
            height: display.height,
            fallback: null,
          };

    if (!isValidResolvedImage(resolved)) {
      // The same contract check `resolveDocumentImages` applies to a
      // third-party resolver, applied to a third-party rasteriser.
      warn(
        "rasterize-failed",
        diagram,
        `the rasteriser returned something OOXML cannot hold for diagram ${diagram} ` +
          `(format "${raster.format}", ${raster.data.length} bytes); the fence was kept`,
      );
      continue;
    }

    const picture = paragraph([
      image(`mermaid:diagram-${diagram}`, {
        alt: info.alt,
        ...(info.caption === null ? {} : { title: info.caption }),
        resolved,
      }),
    ]);

    replacements.set(
      fence,
      wantsCaption && info.caption !== null
        ? [picture, paragraph([text(info.caption, [italic()])])]
        : [picture],
    );
    rendered += 1;
  }

  if (replacements.size === 0) {
    return { document: doc, diagrams: fences.length, rendered, warnings };
  }

  const children = rewriteBlocks(doc.children, replacements);
  return {
    document: children === doc.children ? doc : { ...doc, children },
    diagrams: fences.length,
    rendered,
    warnings,
  };
}
