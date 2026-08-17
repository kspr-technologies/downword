import type { Token } from "markdown-it";

import {
  blockquote,
  codeBlock,
  footnoteDefinition,
  heading,
  htmlBlock,
  list,
  listItem,
  mathBlock,
  nodeText,
  paragraph,
  table,
  tableCell,
  tableRow,
  text,
  thematicBreak,
  type BlockNode,
  type HeadingLevel,
  type InlineNode,
  type ListItemNode,
  type ListKind,
  type ListNode,
  type TableAlignment,
  type TableCellNode,
  type TableNode,
  type TableRowNode,
} from "../model.js";
import type { ParseContext } from "./context.js";
import type { TokenCursor } from "./cursor.js";
import { attrString, parseInlineTokens, readFootnoteIdentity, summarise } from "./inline.js";

/** Maps a heading token's HTML tag to a model level. OOXML stops at 6. */
const HEADING_LEVELS: Readonly<Record<string, HeadingLevel>> = {
  h1: 1,
  h2: 2,
  h3: 3,
  h4: 4,
  h5: 5,
  h6: 6,
};

/** Token types that close a table's structural wrappers and carry no content. */
const TABLE_WRAPPERS: ReadonlySet<string> = new Set([
  "thead_open",
  "thead_close",
  "tbody_open",
  "tbody_close",
  "tfoot_open",
  "tfoot_close",
]);

/**
 * GFM task-list marker at the very start of an item: `[ ]`, `[x]` or `[X]`,
 * followed by whitespace or the end of the item.
 *
 * The `|$` alternative is what `markdown-it-task-lists` gets wrong — it
 * hard-codes `"[x] "` with a trailing space and so misses the bare `- [x]`
 * that LLMs emit for an empty checklist row.
 */
const TASK_MARKER = /^\[( |x|X)\](?:[ \t]|$)/;

/** `markdown-it-task-lists`' synthesised checkbox, if that plugin is injected. */
const TASK_CHECKBOX_HTML = /<input[^>]*\bclass="task-list-item-checkbox"/;

/**
 * Reads a run of block tokens.
 *
 * @param cursor - Positioned on the first token of the run.
 * @param ctx - Slugger + warning sink.
 * @param closer - Token type that ends the run and is consumed, or `null` to
 *   read to the end of the stream.
 * @param observe - Called with every token *before* it is dispatched, while it
 *   is still known to be a direct child of this run. Used for list tightness,
 *   which depends on `paragraph_open.hidden` on the item's own paragraphs and
 *   must not see paragraphs belonging to a nested list.
 */
export function readBlocks(
  cursor: TokenCursor,
  ctx: ParseContext,
  closer: string | null,
  observe?: (token: Token) => void,
): BlockNode[] {
  const blocks: BlockNode[] = [];

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (closer !== null && token.type === closer) {
      cursor.next();
      break;
    }

    observe?.(token);

    const before = cursor.index;
    readBlock(cursor, ctx, blocks);

    /* c8 ignore next -- every readBlock branch consumes; belt and braces. */
    if (cursor.index === before) cursor.skip();
  }

  return blocks;
}

/**
 * Reads exactly one block construct and appends the resulting node(s).
 *
 * Appends rather than returns because `footnote_block_open` expands to *many*
 * `footnoteDefinition` nodes, and because some tokens (a stray
 * `footnote_anchor`) map to none at all.
 *
 * Invariant: always consumes at least one token.
 */
function readBlock(cursor: TokenCursor, ctx: ParseContext, out: BlockNode[]): void {
  const token = cursor.peek();
  if (token === null) return;

  ctx.noteLine(token);

  switch (token.type) {
    case "paragraph_open": {
      cursor.next();
      out.push(paragraph(readInlineUntil(cursor, ctx, "paragraph_close")));
      return;
    }

    case "heading_open": {
      cursor.next();
      const level = HEADING_LEVELS[token.tag];
      if (level === undefined) {
        ctx.warn("unexpected-token", `heading with unknown tag "${token.tag}"`, token);
      }
      const children = readInlineUntil(cursor, ctx, "heading_close");
      const id = ctx.slugger.slug(children.map(nodeText).join(""));
      out.push(heading(level ?? 1, children, { id }));
      return;
    }

    case "bullet_list_open":
    case "ordered_list_open":
      out.push(readList(cursor, ctx, token));
      return;

    case "blockquote_open":
      cursor.next();
      out.push(blockquote(readBlocks(cursor, ctx, "blockquote_close")));
      return;

    case "table_open":
      out.push(readTable(cursor, ctx));
      return;

    case "fence": {
      cursor.next();
      const { lang, meta } = splitInfoString(token.info);
      out.push(codeBlock(token.content, { lang, meta }));
      return;
    }

    case "code_block":
      cursor.next();
      out.push(codeBlock(token.content));
      return;

    case "hr":
      cursor.next();
      out.push(thematicBreak());
      return;

    case "html_block":
      cursor.next();
      ctx.warn("raw-html", `raw HTML block: ${summarise(token.content)}`, token);
      out.push(htmlBlock(token.content));
      return;

    case "math_block":
    case "math_block_eqno":
      cursor.next();
      out.push(mathBlock(token.content));
      return;

    case "footnote_block_open":
      readFootnoteBlock(cursor, ctx, out);
      return;

    // markdown-it-footnote puts the back-link anchor of a definition whose last
    // child is a *block* (a list, say) directly in the block stream.
    case "footnote_anchor":
      cursor.next();
      return;

    default:
      // A stray closer is a *structural* surprise (some plugin unbalanced the
      // stream), not an unrecognised construct — worth distinguishing, because
      // the two have completely different causes.
      if (token.nesting === -1) {
        ctx.warn("unexpected-token", `orphaned "${token.type}" with no matching opener`, token);
      } else {
        ctx.warn("unsupported-token", `unsupported block token "${token.type}"`, token);
      }
      cursor.skip();
      return;
  }
}

/**
 * Reads the inline content between an opening token and `closer`.
 *
 * A paragraph is not simply `paragraph_open, inline, paragraph_close`:
 * markdown-it-footnote splices `footnote_anchor` tokens in before the close,
 * at block level, so the run has to be looped rather than indexed.
 */
function readInlineUntil(cursor: TokenCursor, ctx: ParseContext, closer: string): InlineNode[] {
  const children: InlineNode[] = [];

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (token.type === closer) {
      cursor.next();
      break;
    }

    cursor.next();

    if (token.type === "inline") {
      children.push(...parseInlineTokens(token.children ?? [], ctx));
      continue;
    }
    if (token.type === "footnote_anchor") continue;

    ctx.warn("unexpected-token", `"${token.type}" where inline content was expected`, token);
    if (token.nesting === 1) skipRemainderOf(cursor);
  }

  return children;
}

/** Skips through the closer of a region whose opening token was already eaten. */
function skipRemainderOf(cursor: TokenCursor): void {
  let depth = 1;
  while (depth > 0) {
    const token = cursor.next();
    if (token === null) return;
    depth += token.nesting;
  }
}

/**
 * Reads a bullet or ordered list, including tightness and task state.
 *
 * Three things markdown-it makes you derive yourself:
 *
 * - **Tightness.** There is no flag on the list token. markdown-it marks the
 *   item paragraphs `hidden` instead, so the list is loose iff any *direct*
 *   item paragraph is visible.
 * - **Start ordinal.** Only present as an HTML `start` attribute, and typed
 *   `string | number` because markdown-it 15 stores it as a number.
 * - **Task items.** See {@link extractTaskMarker}.
 *
 * A bullet list containing at least one checkbox item becomes `kind: "task"`.
 * An *ordered* list keeps `kind: "ordered"` even when its items are checkboxes,
 * because the model's `task` kind forces `start` to `null` and losing the
 * numbering of `1. [x] …` would be the worse trade.
 */
function readList(cursor: TokenCursor, ctx: ParseContext, open: Token): ListNode {
  cursor.next();

  const ordered = open.type === "ordered_list_open";
  const closer = ordered ? "ordered_list_close" : "bullet_list_close";
  const items: ListItemNode[] = [];
  let loose = false;
  let anyChecked = false;

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (token.type === closer) {
      cursor.next();
      break;
    }

    if (token.type !== "list_item_open") {
      ctx.warn("unexpected-token", `"${token.type}" directly inside a list`, token);
      cursor.skip();
      continue;
    }

    const item = readListItem(cursor, ctx);
    items.push(item.node);
    if (item.loose) loose = true;
    if (item.node.checked !== null) anyChecked = true;
  }

  const kind: ListKind = ordered ? "ordered" : anyChecked ? "task" : "bullet";
  return list(kind, items, { start: ordered ? readStart(open) : null, tight: !loose });
}

/** Reads one list item plus whether it forces the enclosing list to be loose. */
function readListItem(
  cursor: TokenCursor,
  ctx: ParseContext,
): { readonly node: ListItemNode; readonly loose: boolean } {
  cursor.next();

  let loose = false;
  const blocks = readBlocks(cursor, ctx, "list_item_close", (token) => {
    if (token.type === "paragraph_open" && !token.hidden) loose = true;
  });

  const task = extractTaskMarker(blocks);
  return { node: listItem(task.blocks, { checked: task.checked }), loose };
}

/**
 * Splits a GFM checkbox off the front of a list item.
 *
 * Both spellings are recognised:
 *
 * - the source form, `[ ] `/`[x] ` as the first characters of the item's first
 *   text run (what markdown-it produces on its own); and
 * - `markdown-it-task-lists`' rewritten form, an `<input …>` `html_inline`
 *   token, in case a caller injects that plugin through `plugins`.
 *
 * The marker plus one following space is removed so the item text reads
 * `"done"`, not `" done"`. A leading run that becomes empty is dropped
 * entirely rather than left as a zero-length text node.
 */
function extractTaskMarker(blocks: readonly BlockNode[]): {
  readonly checked: boolean | null;
  readonly blocks: readonly BlockNode[];
} {
  const first = blocks[0];
  if (first === undefined || first.type !== "paragraph") return { checked: null, blocks };

  const leading = first.children[0];
  if (leading === undefined) return { checked: null, blocks };

  if (leading.type === "htmlInline" && TASK_CHECKBOX_HTML.test(leading.value)) {
    const rest = first.children.slice(1);
    const head = rest[0];
    const trimmed =
      head !== undefined && head.type === "text" && head.value.startsWith(" ")
        ? [text(head.value.slice(1), head.marks), ...rest.slice(1)]
        : rest;

    return {
      checked: leading.value.includes("checked"),
      blocks: [paragraph(dropEmptyLeadingText(trimmed)), ...blocks.slice(1)],
    };
  }

  // A marked-up `**[x]**` is not a checkbox; GFM requires the literal marker.
  if (leading.type !== "text" || leading.marks.length > 0) return { checked: null, blocks };

  const match = TASK_MARKER.exec(leading.value);
  const flag = match?.[1];
  if (match === null || flag === undefined) return { checked: null, blocks };

  const remainder = leading.value.slice(match[0].length);
  const rest =
    remainder.length > 0
      ? [text(remainder, leading.marks), ...first.children.slice(1)]
      : first.children.slice(1);

  return {
    checked: flag !== " ",
    blocks: [paragraph(rest), ...blocks.slice(1)],
  };
}

/** Drops a leading zero-length text run left behind by marker removal. */
function dropEmptyLeadingText(children: readonly InlineNode[]): readonly InlineNode[] {
  const head = children[0];
  if (head !== undefined && head.type === "text" && head.value.length === 0) {
    return children.slice(1);
  }
  return children;
}

/** Reads an ordered list's first ordinal, defaulting to `1`. */
function readStart(open: Token): number {
  const raw = attrString(open, "start");
  if (raw === null) return 1;

  const value = Number.parseInt(raw, 10);
  return Number.isSafeInteger(value) ? value : 1;
}

/**
 * Reads a GFM table.
 *
 * Per-column alignment is not on the table token — markdown-it puts it in an
 * inline `style="text-align:…"` attribute on every single cell. The model
 * stores it once on the table, so it is taken from the first row and the
 * duplicates on later rows are discarded.
 *
 * Ragged input needs no handling here: markdown-it already pads short rows with
 * empty cells and truncates long ones to the header's column count.
 */
function readTable(cursor: TokenCursor, ctx: ParseContext): TableNode {
  cursor.next();

  const rows: TableRowNode[] = [];
  let align: readonly TableAlignment[] | null = null;
  let inHeader = false;

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (token.type === "table_close") {
      cursor.next();
      break;
    }

    if (TABLE_WRAPPERS.has(token.type)) {
      if (token.type === "thead_open") inHeader = true;
      if (token.type === "thead_close") inHeader = false;
      cursor.next();
      continue;
    }

    if (token.type !== "tr_open") {
      ctx.warn("unexpected-token", `"${token.type}" directly inside a table`, token);
      cursor.skip();
      continue;
    }

    const row = readTableRow(cursor, ctx, inHeader);
    rows.push(row.node);
    align ??= row.align;
  }

  return table(rows, { align: align ?? [] });
}

/** Reads one table row plus the per-column alignment its cells declare. */
function readTableRow(
  cursor: TokenCursor,
  ctx: ParseContext,
  header: boolean,
): { readonly node: TableRowNode; readonly align: readonly TableAlignment[] } {
  cursor.next();

  const cells: TableCellNode[] = [];
  const align: TableAlignment[] = [];

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (token.type === "tr_close") {
      cursor.next();
      break;
    }

    if (token.type !== "th_open" && token.type !== "td_open") {
      ctx.warn("unexpected-token", `"${token.type}" directly inside a table row`, token);
      cursor.skip();
      continue;
    }

    cursor.next();
    align.push(alignmentOf(token));
    cells.push(
      tableCell(readInlineUntil(cursor, ctx, token.type === "th_open" ? "th_close" : "td_close")),
    );
  }

  return { node: tableRow(cells, { header }), align };
}

/** Reads a cell's alignment out of its `style="text-align:…"` attribute. */
function alignmentOf(cell: Token): TableAlignment {
  const style = attrString(cell, "style");
  if (style === null) return "none";
  if (style.includes("text-align:center")) return "center";
  if (style.includes("text-align:right")) return "right";
  if (style.includes("text-align:left")) return "left";
  return "none";
}

/**
 * Reads the footnote definitions markdown-it-footnote appends to the stream.
 *
 * They stay in `document.children` (that is where they are emitted, and the
 * model keeps traversal uniform); hoisting them into `Document({ footnotes })`
 * is the renderer's job.
 *
 * The 0-based `meta.id` becomes a 1-based `number`, matching both the rendered
 * marker and docx's rule that footnote ids must be `>= 1`.
 */
function readFootnoteBlock(cursor: TokenCursor, ctx: ParseContext, out: BlockNode[]): void {
  cursor.next();

  for (;;) {
    const token = cursor.peek();
    if (token === null) break;
    if (token.type === "footnote_block_close") {
      cursor.next();
      break;
    }

    if (token.type !== "footnote_open") {
      ctx.warn("unexpected-token", `"${token.type}" inside the footnote block`, token);
      cursor.skip();
      continue;
    }

    cursor.next();
    const identity = readFootnoteIdentity(token);
    const children = readBlocks(cursor, ctx, "footnote_close");

    if (identity === null) {
      ctx.warn("unexpected-token", "footnote definition without a usable id", token);
      continue;
    }

    out.push(
      footnoteDefinition(identity.identifier, identity.number, children, {
        label: identity.label,
      }),
    );
  }
}

/**
 * Splits a fence's info string into a language and the rest.
 *
 * ` ```ts title=foo.ts ` is extremely common in LLM output, and the trailing
 * metadata must not end up in `lang` or every highlighter lookup misses.
 */
export function splitInfoString(info: string): {
  readonly lang: string | null;
  readonly meta: string | null;
} {
  const trimmed = info.trim();
  if (trimmed.length === 0) return { lang: null, meta: null };

  const boundary = trimmed.search(/\s/);
  if (boundary < 0) return { lang: trimmed, meta: null };

  const meta = trimmed.slice(boundary).trim();
  return { lang: trimmed.slice(0, boundary), meta: meta.length > 0 ? meta : null };
}
