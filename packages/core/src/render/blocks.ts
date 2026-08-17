/**
 * Block nodes -> `FileChild`s (`Paragraph` | `Table`).
 *
 * The rule this file follows everywhere: **if OOXML can express it as a style
 * property, it lives in `styles.xml`; only positional facts are written onto
 * the element.** In practice that leaves exactly three kinds of direct
 * formatting, each of which genuinely cannot be a style:
 *
 *  - `w:ind` — how far this block has been pushed right by the quotes and list
 *    items above it. A style cannot know its own nesting depth.
 *  - `w:numPr` — which numbering instance and level this list paragraph joins.
 *  - table geometry — column widths derived from the actual content.
 *
 * Everything else (fonts, colours, shading, spacing, borders, keep-with-next,
 * the code block's contextual spacing) is a `<w:pStyle>` reference.
 */

import {
  AlignmentType,
  BookmarkEnd,
  BookmarkStart,
  BorderStyle,
  HeadingLevel as DocxHeadingLevel,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  TableLayoutType,
  VerticalAlignTable,
  WidthType,
  type FileChild,
  type IBordersOptions,
  type IParagraphOptions,
  type IRunPropertiesOptions,
  type ITableBordersOptions,
  type ParagraphChild,
} from "docx";

import {
  assertNever,
  nodeText,
  type BlockNode,
  type CodeBlockNode,
  type HeadingNode,
  type HtmlBlockNode,
  type ListNode,
  type ListItemNode,
  type MathBlockNode,
  type ParagraphNode,
  type TableAlignment,
  type TableCellNode,
  type TableNode,
  type TableRowNode,
} from "../model.js";
import type { BlockContext, BookmarkAnchor, RenderContext } from "./context.js";
import { imageParagraphKind, importOmml, renderInline } from "./inline.js";
import { MAX_LIST_LEVEL } from "./numbering.js";
import { STYLE_IDS } from "./styles.js";
import { safeText } from "./text.js";
import { scopeColor, type Theme } from "./theme.js";
import { clamp } from "./units.js";
import type { HighlightSpan } from "./types.js";

/** Paragraph options minus its content: the "frame" a block is drawn in. */
type ParagraphFrame = Omit<IParagraphOptions, "text" | "children">;

/* -------------------------------------------------------------------------- */
/* Frames                                                                      */
/* -------------------------------------------------------------------------- */

function quoteBorders(theme: Theme): IBordersOptions {
  return {
    left: {
      style: BorderStyle.SINGLE,
      color: theme.colors.quoteBorder,
      size: theme.spacing.quoteBorderSize,
      space: theme.spacing.quoteBorderSpace,
    },
  };
}

/**
 * `<w:bidi/>` for prose, when the document is right-to-left.
 *
 * Exported because `renderDocument` builds three paragraphs of its own (the
 * title block, an empty footnote body, the placeholder in an empty section) and
 * a base direction is a property of the *document*, not of this module.
 */
export function directionFrame(ctx: RenderContext): ParagraphFrame {
  return ctx.options.direction === "rtl" ? { bidirectional: true } : {};
}

/**
 * Clamps a left indent to the text column, warning the first time it has to.
 *
 * Nesting is unbounded in markdown and finite on paper: fourteen `>` levels
 * puts a paragraph 1054 twips past the *right* margin on A4, where Word draws
 * it as a column zero characters wide. `availableWidth` already floors the
 * width it hands the table and image code, but the number written to `w:ind`
 * was unbounded, so past roughly twelve levels the content simply left the
 * page. One `listIndent` of headroom is kept so the deepest legible block still
 * has a gutter to print in.
 */
function indentLeft(ctx: RenderContext, left: number): number {
  const limit = Math.max(0, ctx.page.contentWidth - ctx.theme.spacing.listIndent);
  if (left <= limit) return left;
  ctx.warnOnce(
    "indent-clamped",
    `a block was nested deeply enough to start ${left - limit} twips past the right margin; ` +
      `its indent was clamped to ${limit} twips so it stays on the page`,
  );
  return limit;
}

/**
 * Indent (and, inside a quote, the left rule) for one block.
 *
 * `withQuoteBorder` is false for blocks that either get the rule from their own
 * style (`Quote`) or whose style already owns `w:pBdr` (`HorizontalRule`) —
 * direct border formatting replaces an inherited `w:pBdr` wholesale rather than
 * merging with it, so adding a left rule to a thematic break would delete the
 * rule that *is* the thematic break.
 *
 * `bidi` is false for the blocks whose content is not prose — code, raw HTML,
 * display math. A right-to-left *document* does not make a shell transcript
 * right-to-left, and `<w:bidi/>` on a code block would move its gutter and
 * reverse the reading order of every line.
 */
function frame(
  ctx: RenderContext,
  block: BlockContext,
  withQuoteBorder: boolean,
  bidi = true,
): ParagraphFrame {
  return {
    ...(block.indent > 0 ? { indent: { left: indentLeft(ctx, block.indent) } } : {}),
    ...(withQuoteBorder && block.quoteDepth > 0 ? { border: quoteBorders(ctx.theme) } : {}),
    ...(bidi ? directionFrame(ctx) : {}),
  };
}

/** Text width available to this block, in twips. */
function availableWidth(ctx: RenderContext, block: BlockContext): number {
  return Math.max(720, ctx.page.contentWidth - block.indent);
}

/* -------------------------------------------------------------------------- */
/* Paragraph & heading                                                         */
/* -------------------------------------------------------------------------- */

function renderParagraph(node: ParagraphNode, ctx: RenderContext, block: BlockContext): FileChild {
  const children = renderInline(node.children, ctx, availableWidth(ctx, block));

  const kind = imageParagraphKind(node.children, ctx);
  // Ordinary body text is `Normal` rather than styleless: a paragraph with no
  // `<w:pStyle>` is the one thing Word's Design -> Style Set cannot restyle.
  const style =
    kind === "figure"
      ? STYLE_IDS.figure
      : kind === "placeholder"
        ? STYLE_IDS.imagePlaceholder
        : block.quoteDepth > 0
          ? STYLE_IDS.quote
          : block.inFootnote
            ? STYLE_IDS.footnoteText
            : STYLE_IDS.normal;

  return new Paragraph({
    style,
    ...frame(ctx, block, style !== STYLE_IDS.quote),
    children,
  });
}

const DOCX_HEADINGS = [
  DocxHeadingLevel.HEADING_1,
  DocxHeadingLevel.HEADING_2,
  DocxHeadingLevel.HEADING_3,
  DocxHeadingLevel.HEADING_4,
  DocxHeadingLevel.HEADING_5,
  DocxHeadingLevel.HEADING_6,
] as const;

/**
 * `<w:bookmarkStart>` / `<w:bookmarkEnd>` around a heading's runs.
 *
 * Written by hand instead of with docx's `Bookmark`, which allocates its
 * numeric id from a counter it creates *per instance* and therefore stamps
 * `w:id="1"` on every bookmark in the document (docx 9.7.1,
 * `bookmarkUniqueNumericIdGen()` inside the constructor). `w:id` is
 * `CT_Bookmark`'s identifier and has to be unique within the part: it is what
 * pairs a start with its end, and what Word's Bookmarks dialog, `REF`
 * cross-references and `TOC` fields address a heading by. Ids come from
 * {@link RenderContext.bookmarks}, allocated once in the pre-pass.
 *
 * Neither class is a member of docx's `ParagraphChild` union, but `Paragraph`
 * pushes any non-`Bookmark` child straight onto its root, so both pack
 * correctly — hence the cast, which mirrors `importOmml`'s.
 */
function bookmarked(
  runs: readonly ParagraphChild[],
  anchor: BookmarkAnchor,
): readonly ParagraphChild[] {
  return [
    new BookmarkStart(anchor.name, anchor.id) as unknown as ParagraphChild,
    ...runs,
    new BookmarkEnd(anchor.id) as unknown as ParagraphChild,
  ];
}

function renderHeading(node: HeadingNode, ctx: RenderContext, block: BlockContext): FileChild {
  const runs = renderInline(node.children, ctx, availableWidth(ctx, block));
  const anchor = ctx.bookmarks.get(node.id);

  // The bookmark around the runs is what makes `[jump](#slug)` land here.
  const children: readonly ParagraphChild[] =
    anchor === undefined ? runs : bookmarked(runs, anchor);

  return new Paragraph({
    // `heading` emits <w:pStyle w:val="HeadingN"/> — a real built-in style id,
    // which is what the Navigation pane, TOC fields and PDF tagging read.
    heading: DOCX_HEADINGS[node.level - 1] ?? DocxHeadingLevel.HEADING_6,
    ...frame(ctx, block, true),
    children,
  });
}

/* -------------------------------------------------------------------------- */
/* Lists                                                                       */
/* -------------------------------------------------------------------------- */

function listReference(
  list: ListNode,
  item: ListItemNode,
  ctx: RenderContext,
  level: number,
): string {
  switch (list.kind) {
    case "bullet":
      return ctx.numbering.bullet();
    case "task":
      return ctx.numbering.task(item.checked);
    case "ordered":
      return ctx.numbering.ordered(list.start ?? 1, level);
    default:
      return assertNever(list.kind, "list kind");
  }
}

/**
 * The reference an *ordered* list will join, or `null` for one with no counter.
 *
 * Only ordered lists can continue each other wrongly, and only they therefore
 * need their reference tracked down the tree. Bullets and task lists return
 * `null`, which `renderList` reads as "nothing here can be miscounted, so reuse
 * whatever instance the ancestors are on".
 *
 * An empty list asks for nothing: registering a reference would put an
 * `<w:abstractNum>` in `numbering.xml` that no paragraph ever points at.
 */
function orderedReference(list: ListNode, ctx: RenderContext, level: number): string | null {
  if (list.kind !== "ordered" || list.children.length === 0) return null;
  return ctx.numbering.ordered(list.start ?? 1, level);
}

/**
 * Properties for the one paragraph in a list item that carries the marker.
 *
 * When nothing above has moved the list (`block.indent === 0`) and it fits
 * inside OOXML's nine levels, no `<w:ind>` is written at all and the indent
 * comes from the numbering definition — the cleanest possible output, and the
 * one Word itself produces. Two things force an explicit indent, and then it
 * must restate the hanging indent too, because a direct `<w:ind>` replaces the
 * level's rather than merging with it:
 *
 *  - a quote or an outer item has pushed the list right; or
 *  - the list is nested past level 9, where `w:ilvl` saturates. Without this,
 *    levels 10 and 11 would sit at exactly the same indent as level 9 and the
 *    nesting would vanish from the page. See `renderList`, which also warns.
 */
function listParagraphFrame(
  list: ListNode,
  ctx: RenderContext,
  block: BlockContext,
  reference: string,
): ParagraphFrame {
  const theme = ctx.theme;
  const overflow = Math.max(0, block.listDepth - block.listLevel);
  const shifted = block.indent > 0 || overflow > 0;
  return {
    style: STYLE_IDS.listParagraph,
    numbering: { reference, level: block.listLevel, instance: block.listInstance },
    ...(shifted
      ? {
          indent: {
            left: indentLeft(
              ctx,
              block.indent + theme.spacing.listIndent * (block.listLevel + 1 + overflow),
            ),
            hanging: theme.spacing.listHanging,
          },
        }
      : {}),
    // ListParagraph is tight by default (contextualSpacing). A loose list -
    // which is what most LLM output actually is - turns that off so the items
    // keep their paragraph spacing.
    ...(list.tight
      ? {}
      : { spacing: { after: theme.spacing.listItemAfter }, contextualSpacing: false }),
    ...(block.quoteDepth > 0 ? { border: quoteBorders(theme) } : {}),
    ...directionFrame(ctx),
  };
}

/**
 * The checkbox an *ordered* item has to draw for itself, or `null`.
 *
 * A bullet list whose items carry checkboxes becomes `kind: "task"` and the
 * glyph is the list marker. An ordered one cannot do that — the marker is the
 * number — but the model still records `checked`, and dropping it would turn
 * `1. [x] done` into `1. done` with nothing to show for the marker the author
 * wrote. So it is drawn as the first run of the item instead: `1. ☑ done`.
 */
function taskGlyphRun(
  list: ListNode,
  item: ListItemNode,
  ctx: RenderContext,
  /** False when the glyph is the paragraph's only content, so it needs no gap. */
  followedByText: boolean,
): TextRun | null {
  if (list.kind !== "ordered" || item.checked === null) return null;
  const glyphs = ctx.theme.taskGlyphs;
  const glyph = safeText(item.checked ? glyphs.checked : glyphs.unchecked, ctx);
  return new TextRun({
    text: followedByText ? `${glyph} ` : glyph,
    font: ctx.theme.fonts.symbol,
  });
}

function renderListItem(
  list: ListNode,
  item: ListItemNode,
  ctx: RenderContext,
  block: BlockContext,
): readonly FileChild[] {
  const reference = listReference(list, item, ctx, block.listLevel);
  const markerFrame = listParagraphFrame(list, ctx, block, reference);
  const out: FileChild[] = [];

  const first = item.children[0];
  const startsWithParagraph = first !== undefined && first.type === "paragraph";

  const glyph = taskGlyphRun(list, item, ctx, startsWithParagraph);
  const lead: readonly ParagraphChild[] = glyph === null ? [] : [glyph];

  if (startsWithParagraph) {
    out.push(
      new Paragraph({
        ...markerFrame,
        children: [...lead, ...renderInline(first.children, ctx, availableWidth(ctx, block))],
      }),
    );
  } else {
    // An item whose first block is a nested list or a code fence still needs a
    // marker; give it an empty numbered paragraph rather than losing the bullet.
    out.push(new Paragraph({ ...markerFrame, children: lead }));
  }

  const rest = startsWithParagraph ? item.children.slice(1) : item.children;
  // Continuation blocks line up with the item's text; a nested list keeps the
  // container's indent because its own numbering level already indents it.
  const continuation: BlockContext = {
    ...block,
    indent: block.indent + ctx.theme.spacing.listIndent * (block.listLevel + 1),
  };
  // Only the *first* sub-list of an item can continue this list's numbering:
  // the marker paragraph above it is the `ilvl` N-1 row that restarts level N.
  // A second one has no such row in front of it, so it must not inherit the
  // reference - `renderList` would otherwise let it carry on counting.
  let sublistBlock = block;
  for (const child of rest) {
    if (child.type === "list") {
      out.push(...renderBlocks([child], ctx, sublistBlock));
      sublistBlock = { ...block, listReference: null };
    } else {
      out.push(...renderBlocks([child], ctx, continuation));
    }
  }

  return out;
}

/**
 * Whether anything in this list subtree counts.
 *
 * Only a counter needs restarting, so only a subtree containing an ordered list
 * needs a numbering instance of its own — see `sharedInstance` in
 * `numbering.ts` for why minting one per list regardless is not free.
 */
function containsOrderedList(node: ListNode): boolean {
  if (node.kind === "ordered") return true;
  for (const item of node.children) {
    for (const child of item.children) {
      if (child.type === "list" && containsOrderedList(child)) return true;
    }
  }
  return false;
}

/**
 * Renders one list, and decides which numbering instance it joins.
 *
 * That decision is the whole of list correctness, and it turns on one OOXML
 * rule: **a level restarts when a paragraph at a lower level of the same
 * `numId` appears above it.** So two lists may share an instance only when
 * something between them will actually trigger that restart.
 *
 *  - **Top level.** A list that counts (an ordered list, or a bullet list with
 *    one inside) takes a fresh instance, which docx turns into its own
 *    `<w:num>` with `startOverride`. That is what makes two adjacent `1. 2. 3.`
 *    lists print `1..3` twice instead of `1..6`. One that counts nothing —
 *    bullets and checkboxes all the way down — joins the shared instance, so a
 *    bullet-heavy paste does not mint thousands of concrete numberings (see
 *    `sharedInstance` in `numbering.ts` for what that costs the packer).
 *  - **Nested under the same shape.** `1. / 1. / 2. / 2.` reuses its ancestor's
 *    instance: the enclosing `ilvl` 0 paragraphs sit between the two inner
 *    lists on that very `numId`, so level 1 restarts under each outer item.
 *  - **Nested under a different shape.** A `1. 2.` under a *bullet* parent has
 *    no such row: the bullet paragraphs belong to a different `numId`
 *    entirely, so the inner lists would run `1. 2. 3. 4.` across the whole
 *    document. Those get a fresh instance each, which is what
 *    `block.listReference` exists to detect.
 */
function renderList(node: ListNode, ctx: RenderContext, block: BlockContext): readonly FileChild[] {
  const depth = block.listDepth + 1;
  const level = Math.min(depth, MAX_LIST_LEVEL);
  if (depth > MAX_LIST_LEVEL) {
    ctx.warn(
      "list-depth-truncated",
      `a list nested ${depth + 1} levels deep was flattened onto level 9, the deepest OOXML has; ` +
        `it keeps indenting but shares the ninth level's marker`,
    );
  }

  const reference = orderedReference(node, ctx, level);
  const reuse =
    block.listInstance >= 0 && (reference === null || reference === block.listReference);
  const instance = reuse
    ? block.listInstance
    : containsOrderedList(node)
      ? ctx.numbering.nextInstance()
      : ctx.numbering.sharedInstance();

  const inner: BlockContext = {
    ...block,
    listLevel: level,
    listDepth: depth,
    listInstance: instance,
    // Deliberately `reference`, not `reference ?? block.listReference`: a bullet
    // list between two ordered ones breaks the chain, because its paragraphs are
    // on a different numId and cannot restart the ordered counter either.
    listReference: reference,
  };
  const out: FileChild[] = [];
  for (const item of node.children) {
    out.push(...renderListItem(node, item, ctx, inner));
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Tables                                                                      */
/* -------------------------------------------------------------------------- */

/** Every column is guaranteed room for at least this many characters. */
const FLOOR_COLUMN_CHARS = 4;
/** Below this, extra measured width stops shrinking a column. */
const MIN_COLUMN_CHARS = 8;
/** Above this, extra measured width stops earning a column more space. */
const MAX_COLUMN_CHARS = 48;

/**
 * Splits `total` into integers proportional to `weights` that sum to exactly
 * `total`.
 *
 * Largest-remainder rather than naive rounding, because Word draws a ragged
 * right edge if `sum(gridCol) != tblW`. Ties break on index, so the result is a
 * pure function of the input — a requirement for byte-stable golden tests.
 */
function distribute(total: number, weights: readonly number[]): number[] {
  const count = weights.length;
  if (count === 0) return [];

  const sum = weights.reduce((a, b) => a + b, 0);
  const exact = sum <= 0 ? weights.map(() => total / count) : weights.map((w) => (total * w) / sum);
  const out = exact.map((value) => Math.floor(value));

  let remainder = total - out.reduce((a, b) => a + b, 0);
  const byFraction = exact
    .map((value, index) => ({ index, fraction: value - Math.floor(value) }))
    .sort((a, b) => b.fraction - a.fraction || a.index - b.index);

  for (const { index } of byFraction) {
    if (remainder <= 0) break;
    const current = out[index];
    if (current === undefined) continue;
    out[index] = current + 1;
    remainder -= 1;
  }

  return out;
}

/**
 * Distributes the text width across a table's columns.
 *
 * A flat `total / columns` split wrecks the commonest markdown table — a narrow
 * "Option" column beside a wide "Description" — so columns are measured by
 * their longest cell. But *purely* proportional widths are just as bad in the
 * other direction: a 4-column table with one prose column squeezes the other
 * three below the width of their own header.
 *
 * So each column is first given a hard floor (cell padding plus four
 * characters) and only the space left over is shared out by measured content,
 * clamped to `[8, 48]` characters. Character widths are approximated at half
 * the body font size, which is close enough for the proportional split and
 * never needs font metrics.
 */
function computeColumnWidths(
  table: TableNode,
  columnCount: number,
  total: number,
  theme: Theme,
): readonly number[] {
  // sizes.body is in half-points; an average glyph is about half an em wide.
  // 22 half-points (11 pt) -> ~110 twips per character.
  const charWidth = Math.max(1, Math.round(theme.sizes.body * 5));
  const padding = 2 * theme.spacing.tableCellPaddingX;
  const floor = padding + FLOOR_COLUMN_CHARS * charWidth;

  // Pathologically many columns: nothing to apportion, split evenly.
  if (floor * columnCount >= total) {
    return distribute(
      total,
      Array.from({ length: columnCount }, () => 1),
    );
  }

  const extras: number[] = [];
  for (let column = 0; column < columnCount; column += 1) {
    let longest = 0;
    for (const row of table.children) {
      const cell = row.children[column];
      if (cell === undefined) continue;
      longest = Math.max(longest, nodeText(cell).length);
    }
    const ideal = padding + clamp(longest, MIN_COLUMN_CHARS, MAX_COLUMN_CHARS) * charWidth;
    extras.push(Math.max(0, ideal - floor));
  }

  return distribute(total - floor * columnCount, extras).map((extra) => extra + floor);
}

function tableBorders(theme: Theme): ITableBordersOptions {
  const edge = {
    style: BorderStyle.SINGLE,
    color: theme.colors.tableBorder,
    size: theme.spacing.tableBorderSize,
  };
  return {
    top: edge,
    bottom: edge,
    left: edge,
    right: edge,
    insideHorizontal: edge,
    insideVertical: edge,
  };
}

function cellAlignment(
  alignment: TableAlignment,
): (typeof AlignmentType)[keyof typeof AlignmentType] | null {
  switch (alignment) {
    case "left":
      return AlignmentType.LEFT;
    case "center":
      return AlignmentType.CENTER;
    case "right":
      return AlignmentType.RIGHT;
    case "none":
      return null;
    default:
      return assertNever(alignment, "table alignment");
  }
}

function renderTableCell(
  cell: TableCellNode | null,
  column: number,
  table: TableNode,
  widths: readonly number[],
  header: boolean,
  ctx: RenderContext,
): TableCell {
  const theme = ctx.theme;
  const width = widths[column] ?? 0;
  // GFM's delimiter row is per column, so alignment lives on the table and is
  // applied to the cell's paragraph - `w:jc` has no cell-level equivalent.
  const alignment = cellAlignment(table.align[column] ?? "none");
  const inner = Math.max(360, width - 2 * theme.spacing.tableCellPaddingX);

  return new TableCell({
    children: [
      new Paragraph({
        style: header ? STYLE_IDS.tableHeading : STYLE_IDS.tableText,
        ...(alignment === null ? {} : { alignment }),
        ...directionFrame(ctx),
        children: cell === null ? [] : renderInline(cell.children, ctx, inner),
      }),
    ],
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlignTable.TOP,
    ...(header
      ? {
          shading: {
            type: ShadingType.CLEAR,
            color: "auto",
            fill: theme.colors.tableHeaderBackground,
          },
        }
      : {}),
  });
}

function renderTableRow(
  row: TableRowNode,
  table: TableNode,
  widths: readonly number[],
  columnCount: number,
  ctx: RenderContext,
): TableRow {
  const cells: TableCell[] = [];
  for (let column = 0; column < columnCount; column += 1) {
    // Ragged rows are legal markdown; Word needs every row to fill the grid.
    cells.push(
      renderTableCell(row.children[column] ?? null, column, table, widths, row.header, ctx),
    );
  }
  return new TableRow({
    children: cells,
    // tblHeader repeats the row at the top of every page it spills onto, and
    // cantSplit stops the header itself being broken across the page boundary.
    ...(row.header ? { tableHeader: true, cantSplit: true } : {}),
  });
}

/**
 * The empty paragraph every table is followed by. Not optional.
 *
 * `w:tbl` has no spacing property, so without this a table is welded to the
 * paragraph below it — and, far worse, **two adjacent `w:tbl` elements are one
 * table**. `</w:tbl><w:tbl>` is not two tables to Word: it merges them, so the
 * second table's rows are re-gridded into the first's columns and a
 * two-column/one-column pair comes out visibly corrupt. Two tables separated by
 * a blank line is ordinary LLM output, and §17.4.38's content model is why a
 * block-level element has to sit between them.
 *
 * Emitted after *every* table rather than only between adjacent ones, because
 * the trailing case matters too: a document whose last block is a table
 * otherwise ends `[TBL][SECT]`, which Word repairs by synthesising a paragraph
 * of its own.
 */
function tableSpacer(ctx: RenderContext, block: BlockContext): Paragraph {
  return new Paragraph({ style: STYLE_IDS.tableSpacing, ...frame(ctx, block, true) });
}

function renderTable(
  node: TableNode,
  ctx: RenderContext,
  block: BlockContext,
): readonly FileChild[] {
  const columnCount = node.children.reduce((max, row) => Math.max(max, row.children.length), 0);
  if (columnCount === 0) {
    ctx.warn("table-empty", "table had no cells and was skipped");
    return [];
  }

  const theme = ctx.theme;
  const total = availableWidth(ctx, block);
  const widths = computeColumnWidths(node, columnCount, total, theme);
  const rows = node.children.map((row) => renderTableRow(row, node, widths, columnCount, ctx));

  return [
    new Table({
      rows,
      columnWidths: widths,
      width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
      // Fixed layout makes Word honour the computed widths instead of
      // re-measuring the content and producing a different table every open.
      layout: TableLayoutType.FIXED,
      borders: tableBorders(theme),
      margins: {
        marginUnitType: WidthType.DXA,
        top: theme.spacing.tableCellPaddingY,
        bottom: theme.spacing.tableCellPaddingY,
        left: theme.spacing.tableCellPaddingX,
        right: theme.spacing.tableCellPaddingX,
      },
      ...(block.indent > 0
        ? { indent: { size: indentLeft(ctx, block.indent), type: WidthType.DXA } }
        : {}),
      // Mirrors the column order, which is what an RTL reader expects; the
      // cells' own paragraphs get `<w:bidi/>` from `renderTableCell`.
      ...(ctx.options.direction === "rtl" ? { visuallyRightToLeft: true } : {}),
    }),
    tableSpacer(ctx, block),
  ];
}

/* -------------------------------------------------------------------------- */
/* Code blocks                                                                 */
/* -------------------------------------------------------------------------- */

/** One coloured fragment of a code block, before it is split into lines. */
interface CodePiece {
  readonly text: string;
  readonly props: IRunPropertiesOptions;
}

function isSpanArray(
  value: readonly HighlightSpan[] | Promise<readonly HighlightSpan[]>,
): value is readonly HighlightSpan[] {
  return Array.isArray(value);
}

/** Adapters routinely hand back CSS colours; docx needs bare `RRGGBB`. */
function normalizeColor(color: string): string {
  return color.startsWith("#") ? color.slice(1) : color;
}

/**
 * Colour for one highlight span. **The scope, if there is one, decides alone.**
 *
 * A span that names what it is has handed colour to the theme completely: the
 * palette answers, or the run keeps the `CodeBlock` style's own colour. The
 * span's `color` is read only when there is no scope — an adapter that reports
 * only colours keeps them, which is the other half of the contract in
 * {@link HighlightSpan.scope}.
 *
 * Treating `color` as a *fallback* for unnamed scopes looked harmless and was
 * not: `createHighlighter()` attaches both, so every scope a palette happened
 * to omit was painted in the adapter's ink whatever the theme said. That made
 * `theme: "print"` a no-op for whole languages (all of CSS, whose scopes are
 * dash-separated and so unreachable by `scopeColor`'s prefix walk) and made
 * `codePalette`'s documented replace-not-merge semantics unobservable at render
 * time. The palettes now name every scope `downword/highlight` can emit; a
 * scope neither names is meant to be uncoloured.
 */
function spanColor(span: HighlightSpan, ctx: RenderContext): string | null {
  if (span.scope !== undefined) return scopeColor(ctx.theme, span.scope);
  return span.color === undefined || span.color === "" ? null : normalizeColor(span.color);
}

function spanToPiece(span: HighlightSpan, ctx: RenderContext): CodePiece {
  const color = spanColor(span, ctx);
  return {
    text: safeText(span.text, ctx),
    props: {
      ...(color === null ? {} : { color }),
      ...(span.bold === true ? { bold: true } : {}),
      ...(span.italic === true ? { italics: true } : {}),
    },
  };
}

/**
 * Merges neighbouring pieces that will render as identical runs.
 *
 * The highlighter cannot do all of this itself: it keeps `attr` and `number`
 * apart because it does not know whether the theme will colour them the same,
 * and here we do. Every piece is a `<w:r>`, so this is the difference between
 * a 30 000-line fence emitting one run per token and one run per *change*.
 */
function coalescePieces(pieces: readonly CodePiece[]): readonly CodePiece[] {
  const out: CodePiece[] = [];
  for (const piece of pieces) {
    const last = out[out.length - 1];
    if (
      last !== undefined &&
      last.props.color === piece.props.color &&
      last.props.bold === piece.props.bold &&
      last.props.italics === piece.props.italics
    ) {
      out[out.length - 1] = { text: last.text + piece.text, props: last.props };
    } else {
      out.push(piece);
    }
  }
  return out;
}

/**
 * Turns the model's offset+scope spans into coloured pieces.
 *
 * Offsets are clamped and monotonic so a malformed highlighter cannot drop or
 * duplicate source text: whatever happens, the concatenated pieces still equal
 * `node.value`.
 */
function scopeSpansToPieces(node: CodeBlockNode, ctx: RenderContext): readonly CodePiece[] {
  const source = safeText(node.value, ctx);
  const out: CodePiece[] = [];
  let cursor = 0;

  for (const span of node.highlights ?? []) {
    const start = clamp(span.start, cursor, source.length);
    const end = clamp(span.end, start, source.length);
    if (start > cursor) out.push({ text: source.slice(cursor, start), props: {} });
    if (end > start) {
      const color = scopeColor(ctx.theme, span.scope);
      out.push({ text: source.slice(start, end), props: color === null ? {} : { color } });
    }
    cursor = end;
  }
  if (cursor < source.length) out.push({ text: source.slice(cursor), props: {} });
  return out;
}

/**
 * Turns a {@link Highlighter}'s spans into pieces **that still say what the
 * source said**.
 *
 * A highlighter is a colouring pass. It is also third-party code reached
 * through a documented extension point (Shiki, Prism, Starry Night all plug in
 * here), and nothing on this path used to check that what came back was the
 * same text that went in. An adapter returning `[{ text: "OOPS" }]` replaced
 * the entire listing; one returning the code twice emitted it twice; neither
 * produced a warning. `HighlightSpan`'s docblock stated the invariant as a MUST
 * and only downword's *own* highlighter (`highlight/spans.ts:reconcile`) was
 * held to it.
 *
 * So the same reconciliation runs here, on every span list from every source:
 *
 *  - **equal** — use the spans as they are, which is the overwhelming case and
 *    costs one length comparison plus, at most, one string compare;
 *  - **a proper prefix** — keep the colours and append the remainder unstyled.
 *    That is what a grammar that hit an `illegal` match leaves behind, and
 *    losing the colour on the tail is much better than losing the tail;
 *  - **anything else** — discard the spans and render the block as plain
 *    monospace. Colour is not worth a document that lies.
 *
 * Every deviation raises `highlighter-mismatch`, because an adapter silently
 * having its output thrown away is exactly the kind of thing its author needs
 * to hear about.
 */
function verifiedPieces(
  spans: readonly HighlightSpan[],
  node: CodeBlockNode,
  ctx: RenderContext,
  source: string,
): readonly CodePiece[] {
  let emitted = 0;
  for (const span of spans) emitted += span.text.length;

  const pieces = (): CodePiece[] => spans.map((span) => spanToPiece(span, ctx));
  const plain = (): CodePiece[] => [{ text: source, props: {} }];
  const where = node.lang === null ? "an unlabelled code block" : `a ${node.lang} code block`;

  if (emitted === source.length) {
    let joined = "";
    for (const span of spans) joined += span.text;
    if (joined === source) return coalescePieces(pieces());
    ctx.warn(
      "highlighter-mismatch",
      `the highlighter returned ${emitted} characters for ${where} but not the same ones; ` +
        `its spans were discarded and the block renders as plain monospace`,
    );
    return plain();
  }

  if (emitted < source.length) {
    let joined = "";
    for (const span of spans) joined += span.text;
    if (source.startsWith(joined)) {
      ctx.warn(
        "highlighter-mismatch",
        `the highlighter stopped ${source.length - emitted} characters short of the end of ${where}; ` +
          `the remainder was restored unhighlighted`,
      );
      return coalescePieces([...pieces(), { text: source.slice(joined.length), props: {} }]);
    }
  }

  ctx.warn(
    "highlighter-mismatch",
    `the highlighter returned ${emitted} characters for ${where}, which is ${source.length} long; ` +
      `its spans were discarded and the block renders as plain monospace`,
  );
  return plain();
}

/**
 * Resolves highlighting for one block, in priority order:
 * prepared map -> synchronous highlighter -> model spans -> plain.
 */
function codePieces(node: CodeBlockNode, ctx: RenderContext): readonly CodePiece[] {
  if (node.meta !== null) {
    // ```ts title="app.ts" parses fine and has nowhere to go: OOXML has no
    // caption slot on a paragraph run, and inventing one would be a surprise.
    ctx.warn(
      "code-meta-dropped",
      `the fence metadata ${JSON.stringify(node.meta)} has no OOXML equivalent and was not rendered`,
    );
  }

  // Sanitized once, here, so the text every branch below is compared against is
  // the text that will actually be written.
  const source = safeText(node.value, ctx);

  const prepared = ctx.highlights?.get(node);
  if (prepared !== undefined) return verifiedPieces(prepared, node, ctx, source);

  if (ctx.highlighter !== null) {
    const result = ctx.highlighter.highlight(node.value, node.lang);
    if (isSpanArray(result)) return verifiedPieces(result, node, ctx, source);
    // The promise is discarded, so its rejection must be claimed here: an
    // unhandled rejection terminates the whole process under Node's default
    // `--unhandled-rejections=throw`, which would turn one misbehaving
    // highlighter into a crash long after `renderDocument` has returned.
    void result.catch(() => undefined);
    ctx.warn(
      "highlighter-async",
      "options.highlighter returned a Promise but renderDocument is synchronous; use prepareHighlights() instead",
    );
  }

  if (node.highlights !== null) return scopeSpansToPieces(node, ctx);
  return [{ text: source, props: {} }];
}

/**
 * A literal tab inside `<w:t>` is not a tab stop — Word needs `<w:tab/>`, whose
 * width depends on the paragraph's tab stops rather than on the code's column
 * grid. Expanding to spaces is the only way indentation survives in monospace.
 */
function expandTabs(text: string, tabSize: number): string {
  return tabSize <= 0 ? text : text.replaceAll("\t", " ".repeat(tabSize));
}

function toLines(pieces: readonly CodePiece[], tabSize: number): readonly (readonly CodePiece[])[] {
  const lines: CodePiece[][] = [];
  let current: CodePiece[] = [];

  for (const piece of pieces) {
    const parts = expandTabs(piece.text.replace(/\r\n?/g, "\n"), tabSize).split("\n");
    for (let index = 0; index < parts.length; index += 1) {
      if (index > 0) {
        lines.push(current);
        current = [];
      }
      const part = parts[index] ?? "";
      if (part !== "") current.push({ text: part, props: piece.props });
    }
  }
  lines.push(current);

  // A trailing newline terminates the last line; it does not add a blank one.
  const last = lines[lines.length - 1];
  if (lines.length > 1 && last !== undefined && last.length === 0) lines.pop();
  return lines;
}

function renderCodeBlock(
  node: CodeBlockNode,
  ctx: RenderContext,
  block: BlockContext,
): readonly FileChild[] {
  const lines = toLines(codePieces(node, ctx), ctx.options.tabSize);
  return lines.map(
    (line) =>
      new Paragraph({
        // One paragraph per line, all sharing the CodeBlock style, whose
        // contextualSpacing collapses the gaps between them into nothing.
        style: STYLE_IDS.codeBlock,
        // `bidi: false`: a right-to-left *document* does not make a shell
        // transcript right-to-left, and `<w:bidi/>` here would move the gutter
        // and reverse the reading order of every line.
        ...frame(ctx, block, true, false),
        children: line.map((piece) => new TextRun({ text: piece.text, ...piece.props })),
      }),
  );
}

/* -------------------------------------------------------------------------- */
/* Raw HTML and math                                                           */
/* -------------------------------------------------------------------------- */

function renderHtmlBlock(
  node: HtmlBlockNode,
  ctx: RenderContext,
  block: BlockContext,
): readonly FileChild[] {
  ctx.warn("html-block", `block-level HTML is not converted: ${node.value.split("\n")[0] ?? ""}`);
  if (ctx.options.html === "drop") return [];

  const lines = safeText(node.value, ctx).replace(/\r\n?/g, "\n").split("\n");
  while (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();

  return lines.map(
    (line) =>
      new Paragraph({
        style: STYLE_IDS.htmlBlock,
        ...frame(ctx, block, true, false),
        children: line === "" ? [] : [new TextRun({ text: line })],
      }),
  );
}

function renderMathBlock(
  node: MathBlockNode,
  ctx: RenderContext,
  block: BlockContext,
): readonly FileChild[] {
  if (node.omml !== null) {
    const math = importOmml(node.omml);
    if (math !== null) {
      // mathml2omml only ever emits m:oMath (never m:oMathPara), so a display
      // equation is an inline object in its own paragraph - centre it here.
      return [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          ...frame(ctx, block, true, false),
          children: [math],
        }),
      ];
    }
  }

  ctx.warn("math-unconverted", `display math was not converted to OMML: ${node.value}`);
  return [
    new Paragraph({
      style: STYLE_IDS.codeBlock,
      ...frame(ctx, block, true, false),
      children: [new TextRun({ text: safeText(node.value, ctx) })],
    }),
  ];
}

/* -------------------------------------------------------------------------- */
/* Dispatch                                                                    */
/* -------------------------------------------------------------------------- */

/** Renders a run of block nodes at one position in the tree. */
export function renderBlocks(
  nodes: readonly BlockNode[],
  ctx: RenderContext,
  block: BlockContext,
): FileChild[] {
  const out: FileChild[] = [];

  for (const node of nodes) {
    switch (node.type) {
      case "paragraph":
        out.push(renderParagraph(node, ctx, block));
        break;

      case "heading":
        out.push(renderHeading(node, ctx, block));
        break;

      case "list":
        out.push(...renderList(node, ctx, block));
        break;

      case "table":
        out.push(...renderTable(node, ctx, block));
        break;

      case "codeBlock":
        out.push(...renderCodeBlock(node, ctx, block));
        break;

      case "blockquote":
        out.push(
          ...renderBlocks(node.children, ctx, {
            ...block,
            quoteDepth: block.quoteDepth + 1,
            indent: block.indent + ctx.theme.spacing.quoteIndent,
          }),
        );
        break;

      case "thematicBreak":
        // The rule *is* the style's bottom border, so no quote rule here: a
        // direct w:pBdr would replace the inherited one wholesale.
        out.push(new Paragraph({ style: STYLE_IDS.horizontalRule, ...frame(ctx, block, false) }));
        break;

      case "htmlBlock":
        out.push(...renderHtmlBlock(node, ctx, block));
        break;

      case "mathBlock":
        out.push(...renderMathBlock(node, ctx, block));
        break;

      case "footnoteDefinition":
        // Definitions are hoisted into Document({ footnotes }) before the body
        // is rendered; one reaching here is in the wrong place.
        ctx.warn(
          "footnote-misplaced",
          `footnote definition [^${node.label}] was nested inside another block and was dropped`,
        );
        break;

      default:
        return assertNever(node, "block node");
    }
  }

  return out;
}
