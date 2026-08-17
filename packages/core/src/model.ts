/**
 * downword's intermediate document model.
 *
 * This file is the **neutral contract** between the markdown parser and the
 * OOXML renderer. It has *no imports* — not `docx`, not `markdown-it`, not even
 * other files in this package — and it must stay that way. The parser lowers
 * markdown-it tokens into these nodes; the renderer lifts these nodes into
 * `docx` objects. Neither side may leak its own vocabulary into here.
 *
 * ## Invariants
 *
 * 1. **Every node is a member of one discriminated union**, {@link Node},
 *    discriminated on `type`. There is exactly one place where a node kind is
 *    introduced, so `switch (node.type)` can be checked for exhaustiveness.
 * 2. **No optional properties.** Absence is spelled `null`, never `?`. This
 *    keeps builders total (see requirement below), keeps structural equality
 *    meaningful in tests, and sidesteps `exactOptionalPropertyTypes` friction
 *    at every call site. Builder *arguments* may be optional; node *fields*
 *    never are.
 * 3. **Everything is `readonly`.** Nodes are values. Transformations rebuild
 *    rather than mutate.
 * 4. **Marks compose.** Inline styling is a normalised list of {@link Mark}s
 *    carried by each inline node, not a combinatorial explosion of node kinds.
 *    `bold + italic + inlineCode + link` is four marks on one {@link TextNode}.
 * 5. **Deferred work gets a slot, not a redesign.** Syntax that Phase 2 will
 *    render (footnotes, math) already has node kinds, and resolution results
 *    that a later pass computes (image bytes, highlight spans, OMML) already
 *    have `null`-valued fields. Adding those features must not change this
 *    union.
 *
 * ## Deliberate omissions
 *
 * - **Source positions.** Nodes carry no line/column. Diagnostics that need
 *   them should be emitted by the parser while it still holds the token, and
 *   the nodes that survive a lossy conversion ({@link HtmlBlockNode},
 *   {@link HtmlInlineNode}) carry their raw source instead, which is the more
 *   useful thing to put in a warning.
 * - **Definition lists, colspan/rowspan, admonition callouts.** Not GFM. A
 *   renderer that wants callouts can pattern-match a {@link BlockquoteNode}
 *   whose first paragraph starts with `[!NOTE]` without any model change.
 *
 * @see {@link visit} for traversal, {@link assertNever} for exhaustiveness.
 */

/* -------------------------------------------------------------------------- */
/* Marks                                                                       */
/* -------------------------------------------------------------------------- */

/** Bold / strong emphasis (`**x**`). Renders to `<w:b/>`. */
export interface BoldMark {
  readonly type: "bold";
}

/** Italic / emphasis (`*x*`). Renders to `<w:i/>`. */
export interface ItalicMark {
  readonly type: "italic";
}

/** GFM strikethrough (`~~x~~`). Renders to `<w:strike/>`. */
export interface StrikethroughMark {
  readonly type: "strikethrough";
}

/**
 * Inline code span (`` `x` ``).
 *
 * A mark rather than a node so that ``[**`npm i`**](url)`` needs no special
 * case: it is one text run carrying `link + bold + inlineCode`.
 */
export interface InlineCodeMark {
  readonly type: "inlineCode";
}

/** Superscript (`x^2^`, or raw `<sup>`). Renders to `<w:vertAlign w:val="superscript"/>`. */
export interface SuperscriptMark {
  readonly type: "superscript";
}

/** Subscript (`H~2~O`, or raw `<sub>`). Renders to `<w:vertAlign w:val="subscript"/>`. */
export interface SubscriptMark {
  readonly type: "subscript";
}

/**
 * Highlighted / marked text (`==x==`, or raw `<mark>`).
 *
 * Carries no colour: markdown has none. The renderer picks a
 * `HighlightColor` (Word only supports a fixed palette anyway).
 */
export interface HighlightMark {
  readonly type: "highlight";
}

/**
 * A hyperlink applied to inline content.
 *
 * Modelled as a mark, not a wrapper node, for three reasons: markdown cannot
 * nest links, a mark composes with the styling marks without any tree surgery,
 * and it lets an {@link ImageNode} be linked (`[![badge](i.png)](url)`, which
 * is everywhere in READMEs and LLM output) without a second container kind.
 *
 * A renderer should treat an `href` beginning with `#` as an intra-document
 * anchor (`InternalHyperlink` targeting a {@link HeadingNode.id}) and anything
 * else as an `ExternalHyperlink`. Runs that are adjacent siblings carrying an
 * equal link mark may be coalesced into a single hyperlink.
 */
export interface LinkMark {
  readonly type: "link";
  /** Raw destination exactly as authored; never normalised or resolved here. */
  readonly href: string;
  /** The optional markdown link title (`[a](b "title")`), or `null`. */
  readonly title: string | null;
}

/** Any inline formatting applied to an {@link InlineNode}. */
export type Mark =
  | BoldMark
  | ItalicMark
  | StrikethroughMark
  | InlineCodeMark
  | SuperscriptMark
  | SubscriptMark
  | HighlightMark
  | LinkMark;

/** The `type` discriminant of a {@link Mark}. */
export type MarkType = Mark["type"];

/** Narrows {@link Mark} to the single variant identified by `T`. */
export type MarkOfType<T extends MarkType> = Extract<Mark, { readonly type: T }>;

/* -------------------------------------------------------------------------- */
/* Inline nodes                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Fields shared by every inline node.
 *
 * Marks live on *all* inline nodes, not just text, so that `**a<br>b**`,
 * `**$x$**` and `[![img](i.png)](url)` are representable without special cases.
 * `marks` is always normalised — see {@link normalizeMarks}.
 */
export interface InlineNodeBase {
  /** Normalised, de-duplicated, canonically ordered inline formatting. */
  readonly marks: readonly Mark[];
}

/** A literal run of text. The only node that carries user-visible characters. */
export interface TextNode extends InlineNodeBase {
  readonly type: "text";
  /** Already entity-decoded. Renderers must not unescape it again. */
  readonly value: string;
}

/**
 * An explicit line break (`\` or two trailing spaces).
 *
 * Renders to `<w:br/>` inside the current paragraph — it does **not** start a
 * new paragraph.
 */
export interface HardBreakNode extends InlineNodeBase {
  readonly type: "hardBreak";
}

/**
 * A newline inside a paragraph that markdown collapses to a space.
 *
 * Kept in the model rather than normalised away at parse time so a renderer can
 * choose between "join with a space" (correct for prose) and "preserve" (what
 * some users expect from LLM output). The v1 policy is to emit a single space.
 */
export interface SoftBreakNode extends InlineNodeBase {
  readonly type: "softBreak";
}

/** Raster image formats `docx`'s `ImageRun` accepts directly. */
export type RasterImageFormat = "png" | "jpg" | "gif" | "bmp";

/** Every image format `docx`'s `ImageRun` accepts. */
export type ImageFormat = RasterImageFormat | "svg";

/** Decoded raster bytes plus the intrinsic size needed to compute an EMU extent. */
export interface ResolvedRasterImage {
  readonly format: RasterImageFormat;
  readonly data: Uint8Array;
  /** Intrinsic width in CSS pixels. `docx` multiplies by 9525 to get EMU. */
  readonly width: number;
  /** Intrinsic height in CSS pixels. */
  readonly height: number;
}

/**
 * The output of the image-resolution pass: bytes ready to hand to `ImageRun`.
 *
 * `fallback` is **required by OOXML for SVG** (Word stores the SVG plus a raster
 * twin and non-SVG-aware readers show the twin), and must be `null` for raster
 * formats. A resolver that cannot rasterise an SVG must fail the image rather
 * than produce `{ format: "svg", fallback: null }`.
 */
export interface ResolvedImage {
  readonly format: ImageFormat;
  readonly data: Uint8Array;
  /** Intrinsic width in CSS pixels. */
  readonly width: number;
  /** Intrinsic height in CSS pixels. */
  readonly height: number;
  /** Raster twin; non-`null` if and only if `format === "svg"`. */
  readonly fallback: ResolvedRasterImage | null;
}

/**
 * An image reference (`![alt](src "title")`).
 *
 * `src` is whatever the author wrote — a URL, a relative path or a `data:` URI.
 * Fetching and decoding is a separate, host-dependent, possibly async pass that
 * fills {@link ImageNode.resolved}; the model stays synchronous and pure.
 * `resolved === null` means "not resolved (yet)", and the renderer's policy is
 * to fall back to the alt text.
 */
export interface ImageNode extends InlineNodeBase {
  readonly type: "image";
  readonly src: string;
  /** Alt text; `""` when the author supplied none. */
  readonly alt: string;
  readonly title: string | null;
  readonly resolved: ResolvedImage | null;
}

/**
 * Raw inline HTML (markdown-it's `html_inline`), e.g. `<br>`, `<sub>`, `<img>`.
 *
 * Modelled explicitly so the escape/keep/drop decision is *visible* in the
 * pipeline rather than silently made by the parser. Keeping the node means
 * that policy lives in one place and can be changed without touching the
 * parser.
 *
 * A node of this type only ever exists when {@link ConvertOptions.html} is
 * `"keep"` or `"drop"`, both of which emit a warning. Under the default,
 * `"escape"`, markdown-it never produces an HTML token at all: the angle
 * brackets survive as literal text and nothing reaches this interface.
 */
export interface HtmlInlineNode extends InlineNodeBase {
  readonly type: "htmlInline";
  /** The raw source fragment, tag delimiters included. */
  readonly value: string;
}

/**
 * A footnote marker (`[^1]`) pointing at a {@link FootnoteDefinitionNode}.
 *
 * Phase 2. The node exists now so the union is stable.
 */
export interface FootnoteReferenceNode extends InlineNodeBase {
  readonly type: "footnoteReference";
  /** Normalised key that matches this reference to its definition. */
  readonly identifier: string;
  /** The label as authored (`1`, `note`, …). */
  readonly label: string;
  /**
   * 1-based footnote number, shared by every reference to the same definition.
   * Doubles as the `docx` footnote id, which must be `>= 1` (0 and -1 are
   * reserved for the separator paragraphs).
   */
  readonly number: number;
}

/**
 * Inline TeX (`$x$`).
 *
 * Phase 2. `omml` is the slot for the `temml -> mathml2omml` conversion result
 * so the renderer never has to reach for a TeX engine itself.
 */
export interface MathInlineNode extends InlineNodeBase {
  readonly type: "mathInline";
  /** Raw TeX source, delimiters stripped. */
  readonly value: string;
  /** Pre-converted `<m:oMath>` markup, or `null` if not converted. */
  readonly omml: string | null;
}

/** Anything that can appear inside a paragraph, heading or table cell. */
export type InlineNode =
  | TextNode
  | HardBreakNode
  | SoftBreakNode
  | ImageNode
  | HtmlInlineNode
  | FootnoteReferenceNode
  | MathInlineNode;

/* -------------------------------------------------------------------------- */
/* Block nodes                                                                 */
/* -------------------------------------------------------------------------- */

/** A run of prose. Renders to a single `<w:p>`. */
export interface ParagraphNode {
  readonly type: "paragraph";
  readonly children: readonly InlineNode[];
}

/** ATX/setext heading depth. OOXML only defines `Heading1`…`Heading6`. */
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/**
 * A section heading.
 *
 * `id` is the document-unique anchor slug. It is what `[jump](#id)` resolves
 * against, and Phase 2 emits it as a `Bookmark` so `InternalHyperlink` can
 * target it. The parser is responsible for uniqueness — build ids with a
 * {@link createSlugger} so duplicate headings get `-1`, `-2`, … suffixes the
 * way GitHub does. The {@link heading} builder's default is *not* unique.
 */
export interface HeadingNode {
  readonly type: "heading";
  readonly level: HeadingLevel;
  readonly id: string;
  readonly children: readonly InlineNode[];
}

/** Which numbering scheme a {@link ListNode} uses. */
export type ListKind = "ordered" | "bullet" | "task";

/**
 * An ordered, bulleted or task list. Nests to arbitrary depth: a nested list is
 * a {@link ListNode} inside a {@link ListItemNode}'s `children`.
 */
export interface ListNode {
  readonly type: "list";
  readonly kind: ListKind;
  /**
   * First ordinal for `kind === "ordered"` (markdown's `3.` start syntax);
   * `null` for bullet and task lists.
   *
   * Note for the renderer: `docx` cannot express an arbitrary start on a shared
   * numbering instance — each distinct start value needs its own numbering
   * `reference` whose level 0 declares it.
   */
  readonly start: number | null;
  /**
   * CommonMark tightness. A tight list suppresses the paragraph spacing between
   * items; a loose one keeps it. LLM output is very often loose, and treating
   * it as tight makes the document look wrong, so this is not normalised away.
   */
  readonly tight: boolean;
  readonly children: readonly ListItemNode[];
}

/** Per-column text alignment declared by a GFM table's delimiter row. */
export type TableAlignment = "left" | "center" | "right" | "none";

/**
 * A GFM table.
 *
 * Alignment lives here and only here — cells do not repeat it, so there is a
 * single source of truth. `align[columnIndex]` is the alignment of that column;
 * the array is padded to the widest row by the {@link table} builder, but
 * renderers should still treat a missing entry as `"none"` because ragged rows
 * are legal input.
 */
export interface TableNode {
  readonly type: "table";
  readonly align: readonly TableAlignment[];
  readonly children: readonly TableRowNode[];
}

/**
 * A fenced or indented code block.
 *
 * `highlights` is the slot for a syntax-highlighting pass. It is deliberately a
 * flat span list rather than a tree: `docx` has no nesting inside a run, so the
 * highlighter must flatten anyway, and flat spans map 1:1 onto `TextRun`s.
 */
export interface CodeBlockNode {
  readonly type: "codeBlock";
  /** First word of the info string (`ts`, `python`, …), or `null` when absent. */
  readonly lang: string | null;
  /** Remainder of the info string (`title="a.ts"`, …), or `null`. */
  readonly meta: string | null;
  /** The code exactly as written, including the trailing newline if present. */
  readonly value: string;
  /** Pre-tokenised highlight spans, or `null` if the block was not highlighted. */
  readonly highlights: readonly CodeHighlightSpan[] | null;
}

/**
 * One highlighted region of a {@link CodeBlockNode}'s `value`.
 *
 * Offsets are UTF-16 code-unit indices into `value`, `end` exclusive. Spans in
 * a single block must be **sorted by `start` and non-overlapping**; the
 * highlighter flattens nested scopes before producing them. Gaps between spans
 * are unhighlighted text.
 */
export interface CodeHighlightSpan {
  readonly start: number;
  readonly end: number;
  /**
   * Theme-independent scope name, e.g. `"keyword"`, `"string"`, `"comment"`
   * (highlight.js class names with the `hljs-` prefix removed). Mapping a scope
   * to a colour is the renderer's job.
   */
  readonly scope: string;
}

/** A block quote. Nestable: a `blockquote` may contain another `blockquote`. */
export interface BlockquoteNode {
  readonly type: "blockquote";
  readonly children: readonly BlockNode[];
}

/** A horizontal rule (`---`). Renders as a bottom-bordered empty paragraph. */
export interface ThematicBreakNode {
  readonly type: "thematicBreak";
}

/**
 * Raw block-level HTML (markdown-it's `html_block`), e.g. a `<table>` or
 * `<details>` an LLM emitted instead of markdown.
 *
 * Modelled for the same reason as {@link HtmlInlineNode}: so that "we dropped
 * this" is an explicit, warnable decision rather than a silent gap.
 */
export interface HtmlBlockNode {
  readonly type: "htmlBlock";
  /** The raw source block, verbatim. */
  readonly value: string;
}

/**
 * The body of a footnote (`[^1]: …`).
 *
 * Definitions stay in {@link DocumentNode.children} — conventionally at the end,
 * matching where markdown-it emits them — so that traversal is uniform and
 * there is only one place a node can live. The renderer hoists them into
 * `Document({ footnotes })` keyed by {@link FootnoteDefinitionNode.number} and
 * must **not** emit them into the body flow.
 *
 * Phase 2.
 */
export interface FootnoteDefinitionNode {
  readonly type: "footnoteDefinition";
  /** Normalised key matching {@link FootnoteReferenceNode.identifier}. */
  readonly identifier: string;
  /** The label as authored. */
  readonly label: string;
  /** 1-based footnote number; also the `docx` footnote id. Must be `>= 1`. */
  readonly number: number;
  readonly children: readonly BlockNode[];
}

/**
 * Display TeX (`$$…$$`).
 *
 * Phase 2. `mathml2omml` only ever emits `m:oMath` (never `m:oMathPara`), so a
 * display equation lands inline in its own paragraph — centre it via the
 * paragraph's alignment.
 */
export interface MathBlockNode {
  readonly type: "mathBlock";
  /** Raw TeX source, delimiters stripped. */
  readonly value: string;
  /** Pre-converted `<m:oMath>` markup, or `null` if not converted. */
  readonly omml: string | null;
}

/** Anything that can appear at document level, in a list item or in a quote. */
export type BlockNode =
  | ParagraphNode
  | HeadingNode
  | ListNode
  | TableNode
  | CodeBlockNode
  | BlockquoteNode
  | ThematicBreakNode
  | HtmlBlockNode
  | FootnoteDefinitionNode
  | MathBlockNode;

/* -------------------------------------------------------------------------- */
/* Structural nodes                                                            */
/* -------------------------------------------------------------------------- */

/**
 * One item of a {@link ListNode}.
 *
 * Item content is block-level, which is what makes multi-paragraph items and
 * nested lists work.
 */
export interface ListItemNode {
  readonly type: "listItem";
  /**
   * Task-list state: `true` for `- [x]`, `false` for `- [ ]`, `null` when the
   * item is not a checkbox item (always `null` outside a `kind: "task"` list,
   * and possibly `null` inside one — GFM allows mixed items).
   */
  readonly checked: boolean | null;
  readonly children: readonly BlockNode[];
}

/**
 * One row of a {@link TableNode}.
 *
 * `header` marks the row as a repeating header (`<w:tblHeader/>`). GFM only
 * ever produces one, at index 0, but the flag lives on the row so the model can
 * carry richer input later without a shape change.
 */
export interface TableRowNode {
  readonly type: "tableRow";
  readonly header: boolean;
  readonly children: readonly TableCellNode[];
}

/**
 * One cell of a {@link TableRowNode}. GFM cells hold inline content only.
 *
 * Alignment is *not* stored here — read `table.align[columnIndex]`.
 */
export interface TableCellNode {
  readonly type: "tableCell";
  readonly children: readonly InlineNode[];
}

/** Container nodes that are neither block-level nor inline. */
export type StructuralNode = ListItemNode | TableRowNode | TableCellNode;

/* -------------------------------------------------------------------------- */
/* Document                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Document-level metadata, normally lifted from YAML frontmatter.
 *
 * Total by construction: unknown values are `null`/empty, never absent. The
 * mapped fields are exactly the ones OOXML core properties can hold; anything
 * else survives in {@link DocumentMetadata.custom} and can be emitted as a
 * `customProperties` entry.
 */
export interface DocumentMetadata {
  /** `dc:title`. */
  readonly title: string | null;
  /** `dc:creator`. */
  readonly author: string | null;
  /** `dc:description`. */
  readonly description: string | null;
  /** `cp:keywords`, already split. */
  readonly keywords: readonly string[];
  /** Raw date string as authored; parsing/formatting is the renderer's call. */
  readonly date: string | null;
  /** Every frontmatter key that did not map to a field above, stringified. */
  readonly custom: Readonly<Record<string, string>>;
}

/** The root node. Exactly one per conversion. */
export interface DocumentNode {
  readonly type: "document";
  readonly metadata: DocumentMetadata;
  readonly children: readonly BlockNode[];
}

/* -------------------------------------------------------------------------- */
/* The union                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Every node kind in the model.
 *
 * Note that this shadows the DOM's global `Node` in any file that imports it —
 * which is almost always what you want in parser/renderer code.
 */
export type Node = DocumentNode | BlockNode | StructuralNode | InlineNode;

/** The `type` discriminant of a {@link Node}. */
export type NodeType = Node["type"];

/** Narrows {@link Node} to the single variant identified by `T`. */
export type NodeOfType<T extends NodeType> = Extract<Node, { readonly type: T }>;

/**
 * Runtime list of every {@link NodeType}, in a stable order (root, blocks,
 * structural, inline). Useful for coverage assertions and debug tooling.
 *
 * The `satisfies` clause rejects entries that are not real node types; the
 * companion {@link NODE_TYPES_ARE_EXHAUSTIVE} rejects *missing* ones.
 */
export const NODE_TYPES = [
  "document",
  "paragraph",
  "heading",
  "list",
  "table",
  "codeBlock",
  "blockquote",
  "thematicBreak",
  "htmlBlock",
  "footnoteDefinition",
  "mathBlock",
  "listItem",
  "tableRow",
  "tableCell",
  "text",
  "hardBreak",
  "softBreak",
  "image",
  "htmlInline",
  "footnoteReference",
  "mathInline",
] as const satisfies readonly NodeType[];

/** Node types present in {@link Node} but missing from {@link NODE_TYPES}. */
export type MissingNodeTypes = Exclude<NodeType, (typeof NODE_TYPES)[number]>;

/**
 * Compile-time proof that {@link NODE_TYPES} lists every {@link NodeType}.
 *
 * If a node kind is added to the union without being added to the array, this
 * declaration fails to typecheck with an error that *names the missing type*.
 */
export const NODE_TYPES_ARE_EXHAUSTIVE: [MissingNodeTypes] extends [never]
  ? true
  : MissingNodeTypes = true;

/* -------------------------------------------------------------------------- */
/* Exhaustiveness                                                              */
/* -------------------------------------------------------------------------- */

/**
 * Compile-time exhaustiveness guard.
 *
 * Call it in the `default` branch of a `switch` over a discriminated union. If
 * every variant is handled, `value` is `never` and the call typechecks; if a
 * variant is added, the call becomes a type error at *every* such switch. At
 * runtime it throws, which is the right behaviour for a value that the type
 * system says cannot exist.
 *
 * @param value - The unhandled value, statically `never`.
 * @param context - Short description used in the thrown message.
 * @throws Always.
 *
 * @example
 * ```ts
 * switch (node.type) {
 *   case "text":
 *     return node.value;
 *   default:
 *     return assertNever(node, "inline node");
 * }
 * ```
 */
export function assertNever(value: never, context = "value"): never {
  throw new Error(`downword: unhandled ${context}: ${JSON.stringify(value)}`);
}

/* -------------------------------------------------------------------------- */
/* Mark helpers and builders                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Canonical mark ordering.
 *
 * Declaring it as a full `Record<MarkType, number>` means a new mark type is a
 * compile error here too. `link` sorts first so the outermost construct reads
 * first in fixtures and debug output.
 */
const MARK_ORDER = {
  link: 0,
  bold: 1,
  italic: 2,
  strikethrough: 3,
  inlineCode: 4,
  superscript: 5,
  subscript: 6,
  highlight: 7,
} as const satisfies Readonly<Record<MarkType, number>>;

/** Shared empty mark list; every unmarked node points at this one array. */
const NO_MARKS: readonly Mark[] = Object.freeze([]);

/** Shared empty child list, returned by {@link childrenOf} for leaves. */
const NO_CHILDREN: readonly Node[] = Object.freeze([]);

/**
 * Normalises a mark list: at most one mark per {@link MarkType} (last wins) and
 * a canonical order.
 *
 * Every builder runs its input through this, so two nodes built with the same
 * marks in different orders are structurally equal. That is what makes
 * `toEqual` assertions and byte-stable golden output possible.
 *
 * @param marks - Marks in any order, possibly with duplicates.
 * @returns A new, normalised, frozen-by-convention list.
 */
export function normalizeMarks(marks: readonly Mark[]): readonly Mark[] {
  if (marks.length === 0) return NO_MARKS;

  const byType = new Map<MarkType, Mark>();
  for (const mark of marks) byType.set(mark.type, mark);

  return [...byType.values()].sort((a, b) => MARK_ORDER[a.type] - MARK_ORDER[b.type]);
}

/**
 * Whether `marks` contains a mark of the given type.
 *
 * @param marks - Mark list, normalised or not.
 * @param type - Mark type to look for.
 */
export function hasMark(marks: readonly Mark[], type: MarkType): boolean {
  return marks.some((mark) => mark.type === type);
}

/**
 * Returns the mark of the given type, narrowed to its variant, or `null`.
 *
 * @param marks - Mark list, normalised or not.
 * @param type - Mark type to look for.
 *
 * @example
 * ```ts
 * const href = findMark(node.marks, "link")?.href ?? null;
 * ```
 */
export function findMark<T extends MarkType>(
  marks: readonly Mark[],
  type: T,
): MarkOfType<T> | null {
  for (const mark of marks) {
    if (mark.type === type) return mark as MarkOfType<T>;
  }
  return null;
}

const BOLD: BoldMark = Object.freeze({ type: "bold" });
const ITALIC: ItalicMark = Object.freeze({ type: "italic" });
const STRIKETHROUGH: StrikethroughMark = Object.freeze({ type: "strikethrough" });
const INLINE_CODE: InlineCodeMark = Object.freeze({ type: "inlineCode" });
const SUPERSCRIPT: SuperscriptMark = Object.freeze({ type: "superscript" });
const SUBSCRIPT: SubscriptMark = Object.freeze({ type: "subscript" });
const HIGHLIGHT: HighlightMark = Object.freeze({ type: "highlight" });

/** Builds a {@link BoldMark}. */
export function bold(): BoldMark {
  return BOLD;
}

/** Builds an {@link ItalicMark}. */
export function italic(): ItalicMark {
  return ITALIC;
}

/** Builds a {@link StrikethroughMark}. */
export function strikethrough(): StrikethroughMark {
  return STRIKETHROUGH;
}

/** Builds an {@link InlineCodeMark}. */
export function inlineCode(): InlineCodeMark {
  return INLINE_CODE;
}

/** Builds a {@link SuperscriptMark}. */
export function superscript(): SuperscriptMark {
  return SUPERSCRIPT;
}

/** Builds a {@link SubscriptMark}. */
export function subscript(): SubscriptMark {
  return SUBSCRIPT;
}

/** Builds a {@link HighlightMark}. */
export function highlight(): HighlightMark {
  return HIGHLIGHT;
}

/**
 * Builds a {@link LinkMark}.
 *
 * @param href - Destination exactly as authored; `#anchor` means intra-document.
 * @param title - Markdown link title, or `null`.
 */
export function link(href: string, title: string | null = null): LinkMark {
  return { type: "link", href, title };
}

/* -------------------------------------------------------------------------- */
/* Inline builders                                                             */
/* -------------------------------------------------------------------------- */

/**
 * Builds a {@link TextNode}.
 *
 * @param value - Literal, already entity-decoded text.
 * @param marks - Inline formatting; normalised via {@link normalizeMarks}.
 *
 * @example
 * ```ts
 * text("downword", [bold(), link("https://example.com")]);
 * ```
 */
export function text(value: string, marks: readonly Mark[] = NO_MARKS): TextNode {
  return { type: "text", value, marks: normalizeMarks(marks) };
}

/** Builds a {@link HardBreakNode}. */
export function hardBreak(marks: readonly Mark[] = NO_MARKS): HardBreakNode {
  return { type: "hardBreak", marks: normalizeMarks(marks) };
}

/** Builds a {@link SoftBreakNode}. */
export function softBreak(marks: readonly Mark[] = NO_MARKS): SoftBreakNode {
  return { type: "softBreak", marks: normalizeMarks(marks) };
}

/** Optional fields of the {@link image} builder. */
export interface ImageInit {
  /** Alt text. Defaults to `""`. */
  readonly alt?: string | undefined;
  /** Markdown link title. Defaults to `null`. */
  readonly title?: string | null | undefined;
  /** Pre-resolved bytes. Defaults to `null`. */
  readonly resolved?: ResolvedImage | null | undefined;
  /** Inline formatting (e.g. a {@link LinkMark} for a linked image). */
  readonly marks?: readonly Mark[] | undefined;
}

/**
 * Builds an {@link ImageNode}.
 *
 * @param src - Destination exactly as authored.
 * @param init - Optional alt text, title, resolved bytes and marks.
 */
export function image(src: string, init: ImageInit = {}): ImageNode {
  return {
    type: "image",
    src,
    alt: init.alt ?? "",
    title: init.title ?? null,
    resolved: init.resolved ?? null,
    marks: normalizeMarks(init.marks ?? NO_MARKS),
  };
}

/** Builds an {@link HtmlInlineNode}. */
export function htmlInline(value: string, marks: readonly Mark[] = NO_MARKS): HtmlInlineNode {
  return { type: "htmlInline", value, marks: normalizeMarks(marks) };
}

/** Optional fields of the {@link footnoteReference} builder. */
export interface FootnoteReferenceInit {
  /** Authored label. Defaults to the identifier. */
  readonly label?: string | undefined;
  /** Inline formatting. */
  readonly marks?: readonly Mark[] | undefined;
}

/**
 * Builds a {@link FootnoteReferenceNode}.
 *
 * @param identifier - Normalised key linking the reference to its definition.
 * @param number - 1-based footnote number; also the `docx` footnote id.
 * @param init - Optional authored label and marks.
 */
export function footnoteReference(
  identifier: string,
  number: number,
  init: FootnoteReferenceInit = {},
): FootnoteReferenceNode {
  return {
    type: "footnoteReference",
    identifier,
    label: init.label ?? identifier,
    number,
    marks: normalizeMarks(init.marks ?? NO_MARKS),
  };
}

/** Optional fields of the {@link mathInline} builder. */
export interface MathInlineInit {
  /** Pre-converted OMML. Defaults to `null`. */
  readonly omml?: string | null | undefined;
  /** Inline formatting. */
  readonly marks?: readonly Mark[] | undefined;
}

/**
 * Builds a {@link MathInlineNode}.
 *
 * @param value - Raw TeX, delimiters stripped.
 * @param init - Optional pre-converted OMML and marks.
 */
export function mathInline(value: string, init: MathInlineInit = {}): MathInlineNode {
  return {
    type: "mathInline",
    value,
    omml: init.omml ?? null,
    marks: normalizeMarks(init.marks ?? NO_MARKS),
  };
}

/* -------------------------------------------------------------------------- */
/* Block builders                                                              */
/* -------------------------------------------------------------------------- */

/** Builds a {@link ParagraphNode}. */
export function paragraph(children: readonly InlineNode[]): ParagraphNode {
  return { type: "paragraph", children };
}

/** Optional fields of the {@link heading} builder. */
export interface HeadingInit {
  /**
   * Anchor id. Defaults to `slugify(nodeText(children))`, which is **not**
   * guaranteed unique — the parser should pass an id from a
   * {@link createSlugger} instead.
   */
  readonly id?: string | undefined;
}

/**
 * Builds a {@link HeadingNode}.
 *
 * @param level - 1–6.
 * @param children - Inline content.
 * @param init - Optional explicit anchor id.
 *
 * @example
 * ```ts
 * heading(2, [text("Getting started")]); // id: "getting-started"
 * ```
 */
export function heading(
  level: HeadingLevel,
  children: readonly InlineNode[],
  init: HeadingInit = {},
): HeadingNode {
  const id = init.id ?? slugify(children.map(nodeText).join(""));
  return { type: "heading", level, id, children };
}

/** Optional fields of the {@link list} builder. */
export interface ListInit {
  /**
   * First ordinal. Defaults to `1` for ordered lists and `null` otherwise.
   * Ignored (forced to `null`) for bullet and task lists.
   */
  readonly start?: number | null | undefined;
  /** CommonMark tightness. Defaults to `true`. */
  readonly tight?: boolean | undefined;
}

/**
 * Builds a {@link ListNode}.
 *
 * @param kind - `"ordered"`, `"bullet"` or `"task"`.
 * @param children - Items; nest by putting a `list` inside a `listItem`.
 * @param init - Optional start ordinal and tightness.
 */
export function list(
  kind: ListKind,
  children: readonly ListItemNode[],
  init: ListInit = {},
): ListNode {
  const start = kind === "ordered" ? (init.start ?? 1) : null;
  return { type: "list", kind, start, tight: init.tight ?? true, children };
}

/** Builds a bulleted {@link ListNode}. Shorthand for `list("bullet", …)`. */
export function bulletList(children: readonly ListItemNode[], init: ListInit = {}): ListNode {
  return list("bullet", children, init);
}

/** Builds an ordered {@link ListNode}. Shorthand for `list("ordered", …)`. */
export function orderedList(children: readonly ListItemNode[], init: ListInit = {}): ListNode {
  return list("ordered", children, init);
}

/** Builds a task {@link ListNode}. Shorthand for `list("task", …)`. */
export function taskList(children: readonly ListItemNode[], init: ListInit = {}): ListNode {
  return list("task", children, init);
}

/** Optional fields of the {@link listItem} builder. */
export interface ListItemInit {
  /** Checkbox state; `null` (the default) means "not a task item". */
  readonly checked?: boolean | null | undefined;
}

/**
 * Builds a {@link ListItemNode}.
 *
 * @param children - Block content of the item.
 * @param init - Optional checkbox state.
 */
export function listItem(children: readonly BlockNode[], init: ListItemInit = {}): ListItemNode {
  return { type: "listItem", checked: init.checked ?? null, children };
}

/** Optional fields of the {@link table} builder. */
export interface TableInit {
  /**
   * Per-column alignment. Padded with `"none"` to the widest row; entries past
   * the widest row are kept, so an over-long array is not truncated.
   */
  readonly align?: readonly TableAlignment[] | undefined;
}

/**
 * Builds a {@link TableNode}.
 *
 * @param children - Rows; mark the header row with `tableRow(cells, { header: true })`.
 * @param init - Optional per-column alignment.
 */
export function table(children: readonly TableRowNode[], init: TableInit = {}): TableNode {
  const provided = init.align ?? [];
  const columnCount = children.reduce((widest, row) => Math.max(widest, row.children.length), 0);
  const width = Math.max(columnCount, provided.length);
  const align = Array.from({ length: width }, (_unused, index) => provided[index] ?? "none");
  return { type: "table", align, children };
}

/** Optional fields of the {@link tableRow} builder. */
export interface TableRowInit {
  /** Whether this row repeats as a header. Defaults to `false`. */
  readonly header?: boolean | undefined;
}

/**
 * Builds a {@link TableRowNode}.
 *
 * @param children - Cells.
 * @param init - Optional header flag.
 */
export function tableRow(
  children: readonly TableCellNode[],
  init: TableRowInit = {},
): TableRowNode {
  return { type: "tableRow", header: init.header ?? false, children };
}

/** Builds a header {@link TableRowNode}. Shorthand for `tableRow(…, { header: true })`. */
export function headerRow(children: readonly TableCellNode[]): TableRowNode {
  return tableRow(children, { header: true });
}

/** Builds a {@link TableCellNode}. */
export function tableCell(children: readonly InlineNode[]): TableCellNode {
  return { type: "tableCell", children };
}

/** Optional fields of the {@link codeBlock} builder. */
export interface CodeBlockInit {
  /** Info-string language. Defaults to `null`. */
  readonly lang?: string | null | undefined;
  /** Remainder of the info string. Defaults to `null`. */
  readonly meta?: string | null | undefined;
  /** Pre-tokenised highlight spans. Defaults to `null`. */
  readonly highlights?: readonly CodeHighlightSpan[] | null | undefined;
}

/**
 * Builds a {@link CodeBlockNode}.
 *
 * @param value - The code, verbatim.
 * @param init - Optional language, info-string remainder and highlight spans.
 */
export function codeBlock(value: string, init: CodeBlockInit = {}): CodeBlockNode {
  return {
    type: "codeBlock",
    lang: init.lang ?? null,
    meta: init.meta ?? null,
    value,
    highlights: init.highlights ?? null,
  };
}

/** Builds a {@link BlockquoteNode}. */
export function blockquote(children: readonly BlockNode[]): BlockquoteNode {
  return { type: "blockquote", children };
}

/** Builds a {@link ThematicBreakNode}. */
export function thematicBreak(): ThematicBreakNode {
  return { type: "thematicBreak" };
}

/** Builds an {@link HtmlBlockNode}. */
export function htmlBlock(value: string): HtmlBlockNode {
  return { type: "htmlBlock", value };
}

/** Optional fields of the {@link footnoteDefinition} builder. */
export interface FootnoteDefinitionInit {
  /** Authored label. Defaults to the identifier. */
  readonly label?: string | undefined;
}

/**
 * Builds a {@link FootnoteDefinitionNode}.
 *
 * @param identifier - Normalised key matching its references.
 * @param number - 1-based footnote number; also the `docx` footnote id.
 * @param children - Block content of the note.
 * @param init - Optional authored label.
 */
export function footnoteDefinition(
  identifier: string,
  number: number,
  children: readonly BlockNode[],
  init: FootnoteDefinitionInit = {},
): FootnoteDefinitionNode {
  return {
    type: "footnoteDefinition",
    identifier,
    label: init.label ?? identifier,
    number,
    children,
  };
}

/** Optional fields of the {@link mathBlock} builder. */
export interface MathBlockInit {
  /** Pre-converted OMML. Defaults to `null`. */
  readonly omml?: string | null | undefined;
}

/**
 * Builds a {@link MathBlockNode}.
 *
 * @param value - Raw TeX, delimiters stripped.
 * @param init - Optional pre-converted OMML.
 */
export function mathBlock(value: string, init: MathBlockInit = {}): MathBlockNode {
  return { type: "mathBlock", value, omml: init.omml ?? null };
}

/* -------------------------------------------------------------------------- */
/* Document builders                                                           */
/* -------------------------------------------------------------------------- */

/** Optional fields of the {@link metadata} builder. */
export interface DocumentMetadataInit {
  readonly title?: string | null | undefined;
  readonly author?: string | null | undefined;
  readonly description?: string | null | undefined;
  readonly keywords?: readonly string[] | undefined;
  readonly date?: string | null | undefined;
  readonly custom?: Readonly<Record<string, string>> | undefined;
}

/** Shared empty keyword list. */
const NO_KEYWORDS: readonly string[] = Object.freeze([]);

/** Shared empty custom-property map. */
const NO_CUSTOM: Readonly<Record<string, string>> = Object.freeze({});

/**
 * Builds a total {@link DocumentMetadata} from a partial description.
 *
 * @param init - Any subset of the metadata fields.
 * @returns Metadata with every field present; unknown values are `null`/empty.
 */
export function metadata(init: DocumentMetadataInit = {}): DocumentMetadata {
  return {
    title: init.title ?? null,
    author: init.author ?? null,
    description: init.description ?? null,
    keywords: init.keywords ?? NO_KEYWORDS,
    date: init.date ?? null,
    custom: init.custom ?? NO_CUSTOM,
  };
}

/**
 * Builds a {@link DocumentNode}.
 *
 * Named `doc` rather than `document` to avoid colliding with the DOM global.
 *
 * @param children - Top-level blocks; footnote definitions conventionally last.
 * @param meta - Frontmatter-derived metadata. Defaults to empty metadata.
 */
export function doc(
  children: readonly BlockNode[],
  meta: DocumentMetadata = metadata(),
): DocumentNode {
  return { type: "document", metadata: meta, children };
}

/* -------------------------------------------------------------------------- */
/* Slugs                                                                       */
/* -------------------------------------------------------------------------- */

/** Slug produced for a heading whose text contains no slug-safe characters. */
export const FALLBACK_SLUG = "section";

/**
 * GitHub-flavoured anchor slug: lowercase, punctuation dropped, whitespace runs
 * collapsed to a single hyphen. Letters and digits from any script survive, as
 * do `_` and `-`.
 *
 * Deterministic and dependency-free so that the parser (assigning heading ids)
 * and the renderer (resolving `[x](#anchor)`) always agree.
 *
 * @param input - Plain text, e.g. from {@link nodeText}.
 * @returns The slug, or {@link FALLBACK_SLUG} if nothing survived.
 */
export function slugify(input: string): string {
  const slug = input
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s_-]/gu, "")
    .replace(/\s+/gu, "-");

  return slug.length > 0 ? slug : FALLBACK_SLUG;
}

/** Allocates document-unique slugs. Created by {@link createSlugger}. */
export interface Slugger {
  /**
   * Returns a slug for `input` that has not been returned before, appending
   * `-1`, `-2`, … on collision, exactly as GitHub does.
   */
  readonly slug: (input: string) => string;
  /** Forgets every slug issued so far. */
  readonly reset: () => void;
}

/**
 * Creates a {@link Slugger}.
 *
 * Heading anchors must be unique within a document because Phase 2 emits them
 * as OOXML bookmarks, and duplicate bookmark names make `InternalHyperlink`
 * ambiguous. One slugger per document.
 *
 * @example
 * ```ts
 * const slugger = createSlugger();
 * slugger.slug("Setup"); // "setup"
 * slugger.slug("Setup"); // "setup-1"
 * ```
 */
export function createSlugger(): Slugger {
  const seen = new Map<string, number>();

  return {
    slug: (input: string): string => {
      const base = slugify(input);
      const previous = seen.get(base);

      if (previous === undefined) {
        seen.set(base, 0);
        return base;
      }

      let next = previous + 1;
      let candidate = `${base}-${next}`;
      while (seen.has(candidate)) {
        next += 1;
        candidate = `${base}-${next}`;
      }

      seen.set(base, next);
      seen.set(candidate, 0);
      return candidate;
    },
    reset: (): void => {
      seen.clear();
    },
  };
}

/* -------------------------------------------------------------------------- */
/* Traversal                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * The children of any node, as a uniform list.
 *
 * This is the single place that knows the shape of the tree, and it is an
 * exhaustive switch: adding a node kind to {@link Node} makes this fail to
 * compile.
 *
 * @param node - Any node.
 * @returns Its children, or a shared empty array for leaves.
 */
export function childrenOf(node: Node): readonly Node[] {
  switch (node.type) {
    case "document":
    case "paragraph":
    case "heading":
    case "list":
    case "listItem":
    case "table":
    case "tableRow":
    case "tableCell":
    case "blockquote":
    case "footnoteDefinition":
      return node.children;
    case "codeBlock":
    case "thematicBreak":
    case "htmlBlock":
    case "mathBlock":
    case "text":
    case "hardBreak":
    case "softBreak":
    case "image":
    case "htmlInline":
    case "footnoteReference":
    case "mathInline":
      return NO_CHILDREN;
    default:
      return assertNever(node, "node");
  }
}

/** Coarse classification of a {@link Node}. */
export type NodeCategory = "document" | "block" | "structural" | "inline";

/**
 * Classifies a node as root, block-level, structural or inline.
 *
 * Also an exhaustive switch, so a new node kind must declare where it belongs.
 *
 * @param node - Any node.
 */
export function nodeCategory(node: Node): NodeCategory {
  switch (node.type) {
    case "document":
      return "document";
    case "paragraph":
    case "heading":
    case "list":
    case "table":
    case "codeBlock":
    case "blockquote":
    case "thematicBreak":
    case "htmlBlock":
    case "footnoteDefinition":
    case "mathBlock":
      return "block";
    case "listItem":
    case "tableRow":
    case "tableCell":
      return "structural";
    case "text":
    case "hardBreak":
    case "softBreak":
    case "image":
    case "htmlInline":
    case "footnoteReference":
    case "mathInline":
      return "inline";
    default:
      return assertNever(node, "node");
  }
}

/** Whether `node` is block-level content. */
export function isBlockNode(node: Node): node is BlockNode {
  return nodeCategory(node) === "block";
}

/** Whether `node` is inline content. */
export function isInlineNode(node: Node): node is InlineNode {
  return nodeCategory(node) === "inline";
}

/** Whether `node` is a structural container (list item, table row, table cell). */
export function isStructuralNode(node: Node): node is StructuralNode {
  return nodeCategory(node) === "structural";
}

/** Where a node sits in the tree, as handed to a {@link VisitEnter}. */
export interface VisitContext {
  /** The containing node, or `null` for the traversal root. */
  readonly parent: Node | null;
  /** Index within `parent`'s children, or `-1` for the traversal root. */
  readonly index: number;
  /** Ancestors, outermost first, excluding the node itself. A fresh snapshot. */
  readonly ancestors: readonly Node[];
}

/** Return `"skip"` from a {@link VisitEnter} to not descend into the children. */
export type VisitAction = "skip" | void;

/** Called on the way down. */
export type VisitEnter = (node: Node, context: VisitContext) => VisitAction;

/** Called on the way back up, after the children (or immediately, after `"skip"`). */
export type VisitLeave = (node: Node, context: VisitContext) => void;

/**
 * Depth-first, pre-order walk over a node and its descendants.
 *
 * `leave` always runs for a node that `enter` saw, including when `enter`
 * returned `"skip"` — `"skip"` suppresses the *children*, not the node.
 *
 * @param root - Node to start from; visited first.
 * @param enter - Called before descending. Return `"skip"` to prune.
 * @param leave - Optional; called after the subtree is done.
 *
 * @example
 * ```ts
 * let headings = 0;
 * visit(document, (node) => {
 *   if (node.type === "heading") headings += 1;
 *   if (node.type === "codeBlock") return "skip";
 * });
 * ```
 */
export function visit(root: Node, enter: VisitEnter, leave?: VisitLeave): void {
  const ancestors: Node[] = [];

  const walk = (node: Node, parent: Node | null, index: number): void => {
    const context: VisitContext = { parent, index, ancestors: ancestors.slice() };

    if (enter(node, context) !== "skip") {
      ancestors.push(node);
      for (const [childIndex, child] of childrenOf(node).entries()) {
        walk(child, node, childIndex);
      }
      ancestors.pop();
    }

    leave?.(node, context);
  };

  walk(root, null, -1);
}

/**
 * Collects every descendant (and `root` itself) of a given node type, in
 * document order.
 *
 * @param root - Node to search.
 * @param type - Node type to collect.
 *
 * @example
 * ```ts
 * const notes = nodesOfType(document, "footnoteDefinition");
 * ```
 */
export function nodesOfType<T extends NodeType>(root: Node, type: T): readonly NodeOfType<T>[] {
  const found: NodeOfType<T>[] = [];

  visit(root, (node) => {
    if (node.type === type) found.push(node as NodeOfType<T>);
  });

  return found;
}

/**
 * The plain-text content of a subtree.
 *
 * Text, code and math contribute their characters; both break kinds contribute
 * a newline; raw HTML contributes nothing (its markup is not text) and neither
 * does an image's alt text (which is a description, not content — including it
 * would put "diagram of the pipeline" into a heading's anchor slug).
 *
 * Used for heading anchors and for readable test assertions.
 *
 * @param node - Any node.
 */
export function nodeText(node: Node): string {
  const parts: string[] = [];

  visit(node, (current) => {
    switch (current.type) {
      case "text":
      case "codeBlock":
      case "mathInline":
      case "mathBlock":
        parts.push(current.value);
        break;
      case "hardBreak":
      case "softBreak":
        parts.push("\n");
        break;
      default:
        break;
    }
  });

  return parts.join("");
}
