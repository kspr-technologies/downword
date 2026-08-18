/**
 * The model -> `.docx` renderer.
 *
 * ```ts
 * import { Packer } from "docx";
 * import { renderDocument } from "@ksprtech/downword";
 *
 * const file = renderDocument(model, { onWarning: console.warn });
 * const bytes = await Packer.toArrayBuffer(file);
 * ```
 *
 * `renderDocument` is **synchronous**, and does no I/O and no work of its own
 * that is not a function of the model: everything downword writes —
 * `word/document.xml`, `styles.xml`, `numbering.xml`, `footnotes.xml`, the
 * media parts — is byte-identical for the same model, every time. Everything
 * async lives in {@link prepareHighlights} / {@link prepareImages}, which run
 * first and hand their results in as maps.
 *
 * ## What is *not* reproducible, and why
 *
 * The reproducible unit is the **part**, not the package. Three things in a
 * packed `.docx` come from `docx` (and the JSZip it bundles) rather than from
 * the model, and its public API exposes no way to set any of them. All three
 * are enumerated here because a golden test that diffs whole packages will find
 * them, and because "pure" was previously claimed without this caveat:
 *
 *  - **`docProps/core.xml`** — `dcterms:created` and `dcterms:modified`, read
 *    from the wall clock when the `Document` is constructed (so, inside this
 *    function). Freezing them would be worse than leaving them: a real creation
 *    date is a feature of the file.
 *  - **`word/_rels/document.xml.rels`** and the `r:id` attributes referring to
 *    it in **`word/document.xml`** — one `nanoid` per `ExternalHyperlink`,
 *    minted while packing. Only documents containing an external link have any.
 *  - **the zip container** — `Packer` adds every part without a `date`, so
 *    JSZip stamps each local file header from the clock. DOS timestamps have
 *    two-second resolution, so two packages built more than two seconds apart
 *    differ in *most* of their bytes even when every part inside them is
 *    identical (measured on a minimal document: 5,917 of 9,240). Nothing in
 *    `Packer`'s signature reaches this, which is why neither this function nor
 *    `convert` is documented as byte-reproducible.
 *
 * Nothing else varies: `tests/regressions.test.ts` packs the same model twice
 * and asserts that the only *entries* whose content differs are the first two,
 * and that even those match once the two ambient values are substituted. That
 * per-part comparison is what `tests/helpers/normalize.ts` feeds the golden
 * snapshots, and it is why they are stable while the bytes are not. A caller
 * who needs a hash-stable artefact has to rewrite the archive, not the parts.
 *
 * Output is style-driven rather than formatted inline — see `styles.ts` for why
 * that is the single most important property of the whole package — and lists
 * get real, per-list numbering instances rather than docx's shared built-in
 * one; see `numbering.ts`.
 */

import {
  AlignmentType,
  Document,
  Footer,
  PageNumber,
  Paragraph,
  TableOfContents,
  TextRun,
  type FileChild,
  type ISectionOptions,
  type ParagraphChild,
} from "docx";

import {
  nodesOfType,
  type BlockNode,
  type DocumentMetadata,
  type DocumentNode,
  type FootnoteDefinitionNode,
} from "../model.js";
import { flag, integer, invalid, oneOf, twips } from "../validate.js";
import { directionFrame, renderBlocks } from "./blocks.js";
import {
  createFootnoteIndex,
  ROOT_BLOCK_CONTEXT,
  type BlockContext,
  type BookmarkAnchor,
  type RenderContext,
} from "./context.js";
import { createNumberingRegistry } from "./numbering.js";
import { buildStyles, STYLE_IDS } from "./styles.js";
import { safeText } from "./text.js";
import { resolveTheme } from "./theme.js";
import { reportRenderWarning } from "./types.js";
import type {
  PageInit,
  PageNumbersInit,
  PageNumberSettings,
  PageSettings,
  RenderOptions,
  RenderWarningCode,
  ResolvedRenderOptions,
  TocInit,
  TocSettings,
} from "./types.js";

/* -------------------------------------------------------------------------- */
/* Re-exports: the renderer's public surface                                   */
/* -------------------------------------------------------------------------- */

export { buildStyles, headingStyleId, STYLE_IDS } from "./styles.js";
export {
  createNumberingRegistry,
  LIST_LEVEL_COUNT,
  MAX_LIST_LEVEL,
  NUMBERING_REFERENCES,
  type NumberingRegistry,
} from "./numbering.js";
export { prepareHighlights, prepareImages } from "./prepare.js";
export { RENDER_WARNING_CODES, renderWarningSeverity } from "./types.js";
export {
  DEFAULT_THEME,
  resolveTheme,
  scopeColor,
  type BulletLevel,
  type HeadingSpec,
  type HighlightColorValue,
  type LevelFormatValue,
  type TaskGlyphs,
  type Theme,
  type ThemeColors,
  type ThemeFonts,
  type ThemeInit,
  type ThemeSizes,
  type ThemeSpacing,
} from "./theme.js";
export type {
  FootnotePolicy,
  HighlightMap,
  Highlighter,
  HighlightSpan,
  HtmlPolicy,
  ImageMap,
  ImageResolver,
  PageInit,
  PageNumberFormat,
  PageNumberSettings,
  PageNumbersInit,
  PageSettings,
  PageSize,
  PrepareOptions,
  RenderOptions,
  RenderWarning,
  RenderWarningCode,
  RenderWarningHandler,
  ResolvedRenderOptions,
  SoftBreakPolicy,
  TextDirection,
  TocInit,
  TocSettings,
} from "./types.js";
export type { BookmarkAnchor, FootnoteIndex } from "./context.js";
export type { BaseWarning, WarningSeverity } from "../warnings.js";

/* -------------------------------------------------------------------------- */
/* Option resolution                                                           */
/* -------------------------------------------------------------------------- */

/** Portrait dimensions in twips. */
const PAGE_SIZES = {
  A4: { width: 11906, height: 16838 },
  Letter: { width: 12240, height: 15840 },
} as const;

const DEFAULT_MARGIN = 1440;

/**
 * Word's own ceiling for `w:pgSz`: 22 inches on a side.
 *
 * Not a limit of the format — `ST_TwipsMeasure` is an unsigned long — but the
 * point past which Word clamps the value on open and lays the document out at
 * a size the generator did not choose. A number nobody can print is not a page.
 */
const MAX_PAGE_DIMENSION = 22 * 1440;

/** Below this the text column holds nothing; `availableWidth` floors here too. */
const MIN_CONTENT_EXTENT = 720;

/**
 * Resolves page geometry, refusing one that has no page in it.
 *
 * The clamp to {@link MIN_CONTENT_EXTENT} is for *deep nesting*, which is a
 * property of the content and must degrade rather than fail. A 1x1-twip page,
 * or margins that consume more paper than there is, is a property of the
 * **options** and cannot degrade into anything: every block would be clamped to
 * the same 720-twip column and the result is a document nobody asked for. So
 * that case throws, exactly as `ConvertOptions` already does for a zero page
 * dimension — see `src/validate.ts` for why the two must agree.
 */
function resolvePage(init: PageInit | undefined): PageSettings {
  if (init !== undefined && (typeof init !== "object" || init === null)) {
    throw invalid(`options.page must be an object, got ${JSON.stringify(init)}`);
  }
  const size = init?.size ?? "A4";
  if (typeof size !== "string" && (typeof size !== "object" || size === null)) {
    throw invalid(`options.page.size must be "A4", "Letter" or { width, height }`);
  }
  const dimensions =
    typeof size === "string"
      ? PAGE_SIZES[oneOf(size, ["A4", "Letter"] as const, "page.size", "A4")]
      : {
          width: twips(size.width, "page.size.width", 0),
          height: twips(size.height, "page.size.height", 0),
        };
  if (
    dimensions.width <= 0 ||
    dimensions.height <= 0 ||
    dimensions.width > MAX_PAGE_DIMENSION ||
    dimensions.height > MAX_PAGE_DIMENSION
  ) {
    throw invalid(
      `options.page.size must be between 1 and ${MAX_PAGE_DIMENSION} twips on each side, ` +
        `got ${dimensions.width}x${dimensions.height}`,
    );
  }

  const orientation = oneOf(
    init?.orientation,
    ["portrait", "landscape"] as const,
    "page.orientation",
    "portrait",
  );
  const margin = {
    top: twips(init?.margin?.top, "page.margin.top", DEFAULT_MARGIN),
    right: twips(init?.margin?.right, "page.margin.right", DEFAULT_MARGIN),
    bottom: twips(init?.margin?.bottom, "page.margin.bottom", DEFAULT_MARGIN),
    left: twips(init?.margin?.left, "page.margin.left", DEFAULT_MARGIN),
  };
  // docx swaps w/h itself when orientation is landscape, so the *usable* width
  // is the long edge in that case.
  const across = orientation === "landscape" ? dimensions.height : dimensions.width;
  const down = orientation === "landscape" ? dimensions.width : dimensions.height;

  const contentWidth = across - margin.left - margin.right;
  const contentHeight = down - margin.top - margin.bottom;
  if (contentWidth < MIN_CONTENT_EXTENT || contentHeight < MIN_CONTENT_EXTENT) {
    throw invalid(
      `options.page leaves a text column of ${contentWidth}x${contentHeight} twips; ` +
        `margins must leave at least ${MIN_CONTENT_EXTENT} twips in each direction`,
    );
  }

  return {
    width: dimensions.width,
    height: dimensions.height,
    orientation,
    margin,
    contentWidth,
    contentHeight,
  };
}

/**
 * Largest `tabSize` the renderer accepts.
 *
 * `expandTabs` does `" ".repeat(tabSize)` per tab, so the value is a
 * multiplier on the size of every code block in the document. Unbounded, it is
 * not a wide tab stop: `renderDocument(doc, { tabSize: 1e9 })` on a fence
 * containing a single tab threw a bare `RangeError: Invalid string length` out
 * of the renderer, which is precisely the "document Word quietly refuses to
 * open" failure mode `options.ts` promises never to allow. 64 is four times the
 * widest tab stop anyone writes code with.
 */
const MAX_TAB_SIZE = 64;

/** The heading printed above a `TOC` field when the caller names none. */
const DEFAULT_TOC_TITLE = "Contents";

/** Word's own default TOC range, `TOC \o "1-3"`. */
const DEFAULT_TOC_MIN_LEVEL = 1;
const DEFAULT_TOC_MAX_LEVEL = 3;

/**
 * Widens a `boolean | Init` option into the object form, or `null` for "off".
 *
 * `true` means "on with every default", which is the shape both
 * {@link RenderOptions.toc} and {@link RenderOptions.pageNumbers} accept, and
 * `undefined`/`false` are the same "off" — but a `0`, a string or an array is a
 * caller mistake and has to say so rather than quietly meaning `false`.
 */
function asInit<T extends object>(value: boolean | T | undefined, option: string): T | null {
  if (value === undefined || value === false) return null;
  if (value === true) return {} as T;
  if (typeof value !== "object" || value === null) {
    throw invalid(`options.${option} must be a boolean or an object, got ${JSON.stringify(value)}`);
  }
  return value;
}

function resolveToc(init: boolean | TocInit | undefined): TocSettings | null {
  const source = asInit<TocInit>(init, "toc");
  if (source === null) return null;

  const minLevel = integer(source.minLevel, "toc.minLevel", DEFAULT_TOC_MIN_LEVEL, 1, 6);
  const maxLevel = integer(source.maxLevel, "toc.maxLevel", DEFAULT_TOC_MAX_LEVEL, 1, 6);
  if (minLevel > maxLevel) {
    throw invalid(
      `options.toc.minLevel (${minLevel}) must not be deeper than options.toc.maxLevel (${maxLevel})`,
    );
  }

  const rawTitle = source.title;
  if (rawTitle !== undefined && rawTitle !== null && typeof rawTitle !== "string") {
    throw invalid(`options.toc.title must be a string or null, got ${JSON.stringify(rawTitle)}`);
  }

  return { minLevel, maxLevel, title: rawTitle === undefined ? DEFAULT_TOC_TITLE : rawTitle };
}

function resolvePageNumbers(
  init: boolean | PageNumbersInit | undefined,
): PageNumberSettings | null {
  const source = asInit<PageNumbersInit>(init, "pageNumbers");
  if (source === null) return null;

  return {
    format: oneOf(
      source.format,
      ["number", "page-x", "page-x-of-y"] as const,
      "pageNumbers.format",
      "number",
    ),
    alignment: oneOf(
      source.alignment,
      ["left", "center", "right"] as const,
      "pageNumbers.alignment",
      "center",
    ),
  };
}

function resolveOptions(options: RenderOptions): ResolvedRenderOptions {
  return {
    html: oneOf(options.html, ["raw", "drop"] as const, "html", "raw"),
    softBreak: oneOf(
      options.softBreak,
      ["space", "break", "ignore"] as const,
      "softBreak",
      "space",
    ),
    direction: oneOf(options.direction, ["ltr", "rtl"] as const, "direction", "ltr"),
    tabSize: integer(options.tabSize, "tabSize", 4, 0, MAX_TAB_SIZE),
    titleBlock: flag(options.titleBlock, "titleBlock", true),
    footnotes: flag(options.footnotes, "footnotes", true),
    toc: resolveToc(options.toc),
    pageNumbers: resolvePageNumbers(options.pageNumbers),
  };
}

/* -------------------------------------------------------------------------- */
/* Bookmarks                                                                   */
/* -------------------------------------------------------------------------- */

/** Word truncates bookmark names past 40 characters. */
const MAX_BOOKMARK_LENGTH = 40;

/**
 * Turns a heading slug into a legal Word bookmark name.
 *
 * Word's bookmark names admit only letters, digits and underscores and must
 * start with a letter, so `my-heading` becomes `my_heading`. Because that
 * mapping is lossy (`a-b` and `a_b` collide) the result is deduplicated with a
 * numeric suffix, and *both* the heading and every `[…](#slug)` pointing at it
 * go through the same table — so a collision can never silently retarget a
 * link.
 */
function toBookmarkName(slug: string, used: Set<string>): string {
  const cleaned = slug.replace(/[^A-Za-z0-9_]/g, "_").slice(0, MAX_BOOKMARK_LENGTH);
  const base = /^[A-Za-z]/.test(cleaned) ? cleaned : `_${cleaned}`;

  let candidate = base.slice(0, MAX_BOOKMARK_LENGTH);
  let counter = 2;
  while (used.has(candidate)) {
    const suffix = `_${counter}`;
    candidate = `${base.slice(0, MAX_BOOKMARK_LENGTH - suffix.length)}${suffix}`;
    counter += 1;
  }
  used.add(candidate);
  return candidate;
}

/**
 * The first `w:id` a bookmark may take.
 *
 * `CT_Bookmark/@w:id` is an `ST_DecimalNumber`, so 0 is legal, but Word writes
 * its own from 0 and the value is only ever compared, never ordered — starting
 * at 1 keeps a generated part visually distinguishable from a Word-authored one
 * in a diff without meaning anything to a reader.
 */
const FIRST_BOOKMARK_ID = 1;

/**
 * The first `wp:docPr/@id` a drawing may take.
 *
 * `ST_DrawingElementId` is an `xsd:unsignedInt`; Word numbers its own drawings
 * from 1, and matching that keeps a generated file unremarkable to a reader.
 */
const FIRST_DRAWING_ID = 1;

/**
 * A monotonic id counter, one per `renderDocument` call.
 *
 * Two OOXML id spaces need this and both need it for the *same* reason, which
 * is worth stating once: docx 9.7.1 builds its unique-id generators **inside
 * the constructor** of the component that uses them, so each instance gets a
 * fresh counter and every element it numbers comes out as `1`.
 *
 * - `CT_Bookmark/@w:id` (§17.13.6.2) must be unique within the part — it pairs
 *   a `<w:bookmarkStart>` with its `<w:bookmarkEnd>` and is what Word's
 *   Bookmarks dialog, `REF` cross-references and `TOC` fields address a heading
 *   by. docx's `bookmarkUniqueNumericIdGen()` has the flaw, which is why the
 *   renderer writes `BookmarkStart`/`BookmarkEnd` itself.
 * - `wp:docPr/@id` must be unique across the document. docx's
 *   `docPropertiesUniqueNumericIdGen()` has it too, so two pictures both come
 *   out as `id="1"`; the renderer therefore passes `altText.id` explicitly.
 *
 * Each allocator is deliberately shared as the *same function object* by every
 * site that mints from its space — the heading pre-pass and
 * {@link RenderContext.nextBookmarkId} draw from one counter, not two — so a
 * feature that invents an id cannot collide with one that already existed.
 */
function createIdAllocator(first: number): () => number {
  let next = first;
  return () => {
    const id = next;
    next += 1;
    return id;
  };
}

/**
 * Pre-pass: heading slug -> the anchor it becomes, for the whole document.
 *
 * Done before rendering so that a link appearing *above* its target still
 * resolves — the common "table of contents at the top" shape.
 */
function buildBookmarks(
  doc: DocumentNode,
  allocate: () => number,
): ReadonlyMap<string, BookmarkAnchor> {
  const used = new Set<string>();
  const anchors = new Map<string, BookmarkAnchor>();

  for (const heading of nodesOfType(doc, "heading")) {
    if (heading.id === "" || anchors.has(heading.id)) continue;
    anchors.set(heading.id, { name: toBookmarkName(heading.id, used), id: allocate() });
  }

  return anchors;
}

/* -------------------------------------------------------------------------- */
/* Metadata                                                                    */
/* -------------------------------------------------------------------------- */

interface CoreProperties {
  readonly title?: string;
  readonly creator: string;
  readonly lastModifiedBy: string;
  readonly description?: string;
  readonly keywords?: string;
  readonly customProperties?: readonly { readonly name: string; readonly value: string }[];
}

/**
 * `docProps/core.xml`, from the model's metadata.
 *
 * Every value goes through {@link safeText}: `docProps/core.xml` is a separate
 * XML part with the same `Char` production as `document.xml`, so a single
 * `U+000B` in a title pasted from a PDF makes Word reject the whole package —
 * and the body of the document, which is perfectly well-formed, goes with it.
 * `convert()` already sanitizes metadata in `resolveConvertOptions`; a model a
 * consumer built has been through nothing.
 */
function coreProperties(metadata: DocumentMetadata, ctx: RenderContext): CoreProperties {
  const safe = (value: string): string => safeText(value, ctx);
  const custom = Object.entries(metadata.custom).map(([name, value]) => ({
    name: safe(name),
    value: safe(value),
  }));
  // OOXML core properties have no "date the author wrote this" field that is
  // not also a timestamp Word rewrites, so an authored date becomes a custom
  // property instead of being silently dropped.
  if (metadata.date !== null) custom.push({ name: "Date", value: safe(metadata.date) });

  // docx defaults both name fields to the string "Un-named", which Word then
  // shows as the document's Author. An empty string is falsy to docx's own
  // `if (options.creator)` guard, so it omits the element entirely - and an
  // absent author is the truth for a document nobody claimed.
  const author = metadata.author === null ? "" : safe(metadata.author);

  return {
    ...(metadata.title !== null ? { title: safe(metadata.title) } : {}),
    creator: author,
    lastModifiedBy: author,
    ...(metadata.description !== null ? { description: safe(metadata.description) } : {}),
    ...(metadata.keywords.length > 0 ? { keywords: metadata.keywords.map(safe).join(", ") } : {}),
    ...(custom.length > 0 ? { customProperties: custom } : {}),
  };
}

/* -------------------------------------------------------------------------- */
/* Footnotes                                                                   */
/* -------------------------------------------------------------------------- */

const FOOTNOTE_BLOCK_CONTEXT: BlockContext = { ...ROOT_BLOCK_CONTEXT, inFootnote: true };

/**
 * Hoists `[^1]: …` definitions out of the body and into `word/footnotes.xml`.
 *
 * The model keeps definitions in `document.children` because that is where
 * markdown-it emits them, so they have to be lifted here — and, crucially, kept
 * out of the body flow, or every footnote body would also be printed inline.
 *
 * A footnote part can only hold paragraphs, so a table inside one is dropped
 * with a warning rather than producing a file Word refuses to open.
 *
 * Each body is rendered with its own number *open* (see
 * {@link import("./context.js").FootnoteIndex}), which is what turns a `[^y]`
 * inside footnote `x` into parenthetical text
 * instead of a footnote reference Word cannot draw, and what stops two notes
 * that cite each other from recursing.
 *
 * Only notes the body actually referenced are written, which makes the two
 * parts a matched pair in **both** directions: no reference points at a note
 * that is not there, and no note sits there with no reference pointing at it.
 * The second half matters less for validity and more for honesty — Word draws a
 * footnote only where a marker is, and silently drops an unreferenced one on the
 * next save, so writing one would be pretending the text survived.
 *
 * Returns `{}` when {@link RenderOptions.footnotes} is off: no marker became a
 * reference, because every one of them spliced its note into the sentence that
 * cited it instead.
 */
function buildFootnotes(
  definitions: readonly FootnoteDefinitionNode[],
  ctx: RenderContext,
): Record<string, { readonly children: readonly Paragraph[] }> {
  const footnotes: Record<string, { readonly children: readonly Paragraph[] }> = {};

  for (const definition of definitions) {
    // Duplicate numbers cannot both be kept — `Document({ footnotes })` is a
    // record keyed by id, so the second could only overwrite the first — and
    // the index already decided which one the references were resolved against.
    if (ctx.footnotes.byNumber.get(definition.number) !== definition) {
      ctx.warn(
        "footnote-content-dropped",
        `footnote [^${definition.label}] claims number ${definition.number}, which another ` +
          `definition already has; a footnote id can hold one note, so this one was dropped`,
      );
      continue;
    }
    if (!ctx.footnotes.wasReferenced(definition.number)) continue;

    // Opened while its own body renders, so a `[^y]` inside it is spliced into
    // the text rather than becoming a footnote inside a footnote, and so two
    // notes citing each other terminate.
    const opened = ctx.footnotes.open(definition.number);
    let rendered;
    try {
      rendered = renderBlocks(definition.children, ctx, FOOTNOTE_BLOCK_CONTEXT);
    } finally {
      if (opened) ctx.footnotes.close(definition.number);
    }
    const paragraphs = rendered.filter((child): child is Paragraph => child instanceof Paragraph);

    if (paragraphs.length !== rendered.length) {
      ctx.warn(
        "footnote-content-dropped",
        `footnote [^${definition.label}] contained a block a footnote cannot hold (a table); it was dropped`,
      );
    }
    if (paragraphs.length === 0) {
      paragraphs.push(new Paragraph({ style: STYLE_IDS.footnoteText, ...directionFrame(ctx) }));
    }

    // The key is the docx footnote id; ids 0 and -1 are reserved for the
    // separator paragraphs, which is why the model guarantees number >= 1.
    footnotes[String(definition.number)] = { children: paragraphs };
  }

  // Second pass, because the first is what does the inlining: a note cited only
  // from *inside* another note has its words spliced in there, and is not lost
  // — but which notes those are is only known once every body has rendered.
  for (const definition of definitions) {
    if (ctx.footnotes.wasUsed(definition.number)) continue;
    if (ctx.footnotes.byNumber.get(definition.number) !== definition) continue;
    ctx.warn(
      "footnote-unreferenced",
      `footnote [^${definition.label}] is never referenced, so nothing in the document leads to ` +
        `it and its text was dropped; add a [^${definition.label}] marker, or delete the definition`,
    );
  }

  return footnotes;
}

/* -------------------------------------------------------------------------- */
/* Table of contents                                                           */
/* -------------------------------------------------------------------------- */

/**
 * The `TOC` field, and the heading above it.
 *
 * What OOXML stores is the **instruction**, not the entries:
 * `TOC \o "1-3" \h \z \u` tells the word processor to collect every paragraph
 * whose style declares outline level 1 to 3, hyperlink each entry to it, and
 * hide the page numbers in web layout. `\u` adds paragraphs that carry an
 * outline level directly rather than through their style. The entries
 * themselves are computed on update, so downword emits the field marked dirty
 * (`<w:fldChar w:fldCharType="begin" w:dirty="true"/>`) and sets
 * `<w:updateFields/>` in `word/settings.xml`, which is the whole of what the
 * format lets a generator do. A reader still has to act on it — hence the
 * `toc-needs-update` warning, whose message says how.
 *
 * The heading is `TOCHeading`, whose `<w:outlineLvl w:val="9"/>` is exactly
 * what stops the word "Contents" becoming the first line of the contents.
 */
function buildToc(toc: TocSettings, ctx: RenderContext): readonly FileChild[] {
  ctx.warn(
    "toc-needs-update",
    `a table-of-contents field for heading levels ${toc.minLevel}-${toc.maxLevel} was emitted; ` +
      `OOXML stores the field, not its entries, so it is empty until the reader updates it ` +
      `(Word: answer yes to the prompt on open, or right-click the field -> Update Field / F9; ` +
      `LibreOffice: Tools -> Update -> Indexes and Tables). Readers that do not run fields at all ` +
      `— Google Docs, Pages, Quick Look, most converters — will show nothing there`,
  );

  const out: FileChild[] = [];
  if (toc.title !== null && toc.title !== "") {
    out.push(
      new Paragraph({
        style: STYLE_IDS.tocHeading,
        ...directionFrame(ctx),
        children: [
          new TextRun({
            text: safeText(toc.title, ctx),
            ...(ctx.options.direction === "rtl" ? { rightToLeft: true } : {}),
          }),
        ],
      }),
    );
  }

  out.push(
    new TableOfContents(safeText(toc.title ?? DEFAULT_TOC_TITLE, ctx), {
      headingStyleRange: `${toc.minLevel}-${toc.maxLevel}`,
      hyperlink: true,
      useAppliedParagraphOutlineLevel: true,
      hideTabAndPageNumbersInWebView: true,
    }),
  );

  return out;
}

/* -------------------------------------------------------------------------- */
/* Page numbers                                                                */
/* -------------------------------------------------------------------------- */

const FOOTER_ALIGNMENT = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
} as const satisfies Readonly<Record<PageNumberSettings["alignment"], string>>;

/**
 * The footer carrying the page number.
 *
 * `PAGE` and `NUMPAGES` are ordinary fields, but unlike `TOC` they are computed
 * during **layout** rather than on an explicit update: every reader that
 * paginates at all fills them in, so these need no `updateFields` and no
 * warning. `PageNumber.CURRENT` / `PageNumber.TOTAL_PAGES` are docx's tokens for
 * the two, and only mean anything inside a run's `children`.
 *
 * The wording is English, and deliberately so rather than accidentally: a
 * caller who needs another language can leave `format: "number"`, which has no
 * words in it at all.
 */
function buildPageNumberFooter(settings: PageNumberSettings, ctx: RenderContext): Footer {
  const parts: readonly (string | (typeof PageNumber)[keyof typeof PageNumber])[] =
    settings.format === "number"
      ? [PageNumber.CURRENT]
      : settings.format === "page-x"
        ? ["Page ", PageNumber.CURRENT]
        : ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES];

  const run: ParagraphChild = new TextRun({
    children: [...parts],
    ...(ctx.options.direction === "rtl" ? { rightToLeft: true } : {}),
  });

  return new Footer({
    children: [
      new Paragraph({
        style: STYLE_IDS.footer,
        alignment: FOOTER_ALIGNMENT[settings.alignment],
        ...directionFrame(ctx),
        children: [run],
      }),
    ],
  });
}

/* -------------------------------------------------------------------------- */
/* renderDocument                                                              */
/* -------------------------------------------------------------------------- */

/**
 * Renders a model document into a `docx` `Document` (the `File` object).
 *
 * Pack it with `Packer.toArrayBuffer` / `toBlob` / `toBuffer`.
 */
export function renderDocument(doc: DocumentNode, options: RenderOptions = {}): Document {
  if (typeof options !== "object" || options === null) {
    throw invalid(`options must be an object, got ${JSON.stringify(options)}`);
  }

  const theme = resolveTheme(options.theme);
  const page = resolvePage(options.page);
  const resolved = resolveOptions(options);
  const onWarning = options.onWarning;

  const nextBookmarkId = createIdAllocator(FIRST_BOOKMARK_ID);
  const nextDrawingId = createIdAllocator(FIRST_DRAWING_ID);
  // Degradations that are a property of the *document* rather than of one node
  // report once: a 30-level list would otherwise raise twenty identical
  // `indent-clamped` notices and drown everything else in the stream.
  const reported = new Set<RenderWarningCode>();

  const body: BlockNode[] = [];
  const definitions: FootnoteDefinitionNode[] = [];
  for (const child of doc.children) {
    if (child.type === "footnoteDefinition") definitions.push(child);
    else body.push(child);
  }

  const ctx: RenderContext = {
    theme,
    options: resolved,
    page,
    numbering: createNumberingRegistry(theme),
    bookmarks: buildBookmarks(doc, nextBookmarkId),
    nextBookmarkId,
    nextDrawingId,
    // Built before anything is rendered, for the same reason the bookmark table
    // is: the first `[^x]` in the document has to know whether the definition
    // that answers it exists, and it may well be written below.
    footnotes: createFootnoteIndex(definitions),
    highlights: options.highlights ?? null,
    highlighter: options.highlighter ?? null,
    images: options.images ?? null,
    warn: (code: RenderWarningCode, message: string) => {
      reportRenderWarning(onWarning, code, message);
    },
    warnOnce: (code: RenderWarningCode, message: string) => {
      if (reported.has(code)) return;
      reported.add(code);
      reportRenderWarning(onWarning, code, message);
    },
  };

  const children = renderBlocks(body, ctx, ROOT_BLOCK_CONTEXT);

  // Front matter, innermost first: title, then contents heading, then the field.
  if (resolved.toc !== null) children.unshift(...buildToc(resolved.toc, ctx));

  if (resolved.titleBlock && doc.metadata.title !== null) {
    children.unshift(
      new Paragraph({
        style: STYLE_IDS.title,
        ...directionFrame(ctx),
        children: [
          new TextRun({
            text: safeText(doc.metadata.title, ctx),
            ...(resolved.direction === "rtl" ? { rightToLeft: true } : {}),
          }),
        ],
      }),
    );
  }

  // Footnote bodies can contain lists, so they must be rendered before the
  // numbering registry is frozen.
  const footnotes = buildFootnotes(definitions, ctx);
  const numbering = ctx.numbering.build();

  const section: ISectionOptions = {
    properties: {
      page: {
        size: { width: page.width, height: page.height, orientation: page.orientation },
        margin: page.margin,
      },
    },
    ...(resolved.pageNumbers === null
      ? {}
      : { footers: { default: buildPageNumberFooter(resolved.pageNumbers, ctx) } }),
    // A section with no children produces a body Word treats as damaged.
    children:
      children.length > 0
        ? children
        : [new Paragraph({ style: STYLE_IDS.normal, ...directionFrame(ctx) })],
  };

  return new Document({
    sections: [section],
    styles: buildStyles(theme),
    ...(numbering === null ? {} : { numbering }),
    // `<w:updateFields/>`: makes Word offer to fill the TOC in on open. Only
    // set when there is a field that needs it — it prompts the reader, and a
    // document with nothing to update must not.
    ...(resolved.toc === null ? {} : { features: { updateFields: true } }),
    ...(Object.keys(footnotes).length > 0 ? { footnotes } : {}),
    ...coreProperties(doc.metadata, ctx),
    // `dc:subject` has no model field to come from; see RenderOptions.subject.
    ...(options.subject === undefined ? {} : { subject: safeText(options.subject, ctx) }),
  });
}
