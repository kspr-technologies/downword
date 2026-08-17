/**
 * The renderer's public contract: extension points and options.
 *
 * The two extension points ({@link Highlighter} and {@link ImageResolver}) are
 * *asynchronous*, because syntax highlighting and image fetching both are. The
 * renderer itself is deliberately **synchronous**: `renderDocument()` never
 * awaits anything, so it can run inside a `requestIdleCallback`, a worker, or a
 * server handler without colouring the call site.
 *
 * The two are reconciled by a prepare-then-render split. `prepareHighlights()`
 * and `prepareImages()` walk the model, do the async work once, and hand back a
 * plain `Map` keyed by node identity; `renderDocument()` only ever reads from
 * those maps. Adapters therefore implement an async interface and the renderer
 * stays pure.
 */

import type { CodeBlockNode, ImageNode, ResolvedImage } from "../model.js";
import type { BaseWarning, WarningSeverity } from "../warnings.js";
import type { ThemeInit } from "./theme.js";

/* -------------------------------------------------------------------------- */
/* Extension point: syntax highlighting                                        */
/* -------------------------------------------------------------------------- */

/**
 * One coloured run of a code block.
 *
 * A flat list of spans, concatenated, must reproduce the block's source exactly
 * — including newlines, which the renderer splits on to build one paragraph per
 * line. Spans are *not* offsets into the source: an adapter that already has
 * substrings (every highlighter does) hands them over unchanged.
 *
 * **That requirement is checked, not trusted.** A highlighter is a colouring
 * pass; it must not be able to change what the document *says*. The renderer
 * compares the concatenated spans against the block's source and, if they
 * differ, restores the source — appending the missing tail unstyled when the
 * spans are a proper prefix (what a grammar that hit an `illegal` match leaves
 * behind), and discarding the spans entirely otherwise. Either way a
 * `highlighter-mismatch` warning is raised and no character is lost, added or
 * substituted.
 *
 * Every optional field is `| undefined` so adapters compiled with
 * `exactOptionalPropertyTypes` can write `{ text, color: undefined }`.
 */
export interface HighlightSpan {
  /** The literal source text of this span. Never entity-encoded. */
  readonly text: string;
  /**
   * `RRGGBB` without a leading `#`, or omitted for the default code colour.
   *
   * Read **only when the span reports no {@link HighlightSpan.scope}**: a span
   * that names what it is has handed colour to the theme, and mixing the two
   * would let an adapter's ink override the palette for every scope the palette
   * happens not to name.
   */
  readonly color?: string | undefined;
  readonly bold?: boolean | undefined;
  readonly italic?: boolean | undefined;
  /**
   * What this span *is* — `"keyword"`, `"string"`, `"title.function"` — rather
   * than what colour it should be.
   *
   * Reporting it is what lets `theme` mean something for a highlighted block:
   * the renderer resolves the scope against
   * {@link import("./theme.js").Theme.codePalette} (walking `a.b.c` -> `a.b` ->
   * `a`) and uses that colour, or leaves the run at the `CodeBlock` style's own
   * colour when the palette does not name it. {@link HighlightSpan.color} is
   * **not** consulted for a span that reports a scope — if it were, an
   * adapter's own inks would quietly win back every scope the palette omits,
   * and a whole language could stay unthemeable.
   *
   * So: **an adapter that reports scopes is themeable; one that reports only
   * colours keeps them.** `createHighlighter` reports scopes unless you hand it
   * your own `scopeStyles`, which is how a caller takes the palette back.
   *
   * The names are highlight.js's (`hljs-` already stripped), because that is
   * the vocabulary `theme.codePalette` is keyed by; an adapter for a different
   * engine should map onto them or omit the field.
   */
  readonly scope?: string | undefined;
}

/**
 * Turns source code into {@link HighlightSpan}s.
 *
 * `lang` is the fenced block's info string, or `null` for an indented block or
 * a bare fence. An implementation that does not know the language should return
 * a single span containing the whole input rather than throwing.
 */
export interface Highlighter {
  highlight(
    code: string,
    lang: string | null,
  ): readonly HighlightSpan[] | Promise<readonly HighlightSpan[]>;
}

/** Highlight spans keyed by the {@link CodeBlockNode} they belong to. */
export type HighlightMap = ReadonlyMap<CodeBlockNode, readonly HighlightSpan[]>;

/* -------------------------------------------------------------------------- */
/* Extension point: image resolution                                           */
/* -------------------------------------------------------------------------- */

/**
 * Turns an `![alt](src)` destination into bytes the renderer can embed.
 *
 * `src` is exactly what the author wrote: an absolute URL, a relative path or a
 * `data:` URI. Returning `null` means "could not resolve" and makes the
 * renderer emit a placeholder plus a warning; it must never throw for a merely
 * missing image.
 *
 * Note the SVG contract from {@link ResolvedImage}: a resolver that returns
 * `format: "svg"` **must** also supply a raster `fallback`, because OOXML
 * stores an SVG as an SVG *plus* a raster twin.
 */
export interface ImageResolver {
  resolve(src: string): Promise<ResolvedImage | null>;
}

/** Resolved images keyed by the {@link ImageNode} they belong to. */
export type ImageMap = ReadonlyMap<ImageNode, ResolvedImage>;

/* -------------------------------------------------------------------------- */
/* Warnings                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Machine-readable reason a {@link RenderWarning} was raised.
 *
 * Covers the whole render stage, which is the renderer *plus* its two async
 * preludes ({@link import("./prepare.js").prepareHighlights} and
 * {@link import("./prepare.js").prepareImages}) — they degrade the same
 * document and there is no reason for a host to learn two vocabularies.
 */
export type RenderWarningCode =
  /** An `![img]()` had no bytes; a placeholder was emitted instead. */
  | "image-unresolved"
  /** Bytes were present but unusable (zero-sized, or an SVG with no fallback). */
  | "image-invalid"
  /** An {@link ImageResolver} threw or rejected; that picture has no bytes. */
  | "image-resolver-failed"
  /** Raw block-level HTML was encountered; see {@link RenderOptions.html}. */
  | "html-block"
  /** Raw inline HTML was encountered; see {@link RenderOptions.html}. */
  | "html-inline"
  /** `$…$` / `$$…$$` reached the renderer without a converted OMML payload. */
  | "math-unconverted"
  /** A `[link](#anchor)` pointed at a heading id that does not exist. */
  | "link-unresolved"
  /** A `[text](url "title")` title was dropped; docx cannot write `w:tooltip`. */
  | "link-title-dropped"
  /** A fence's info-string metadata (` ```ts title="a.ts" `) has nowhere to go. */
  | "code-meta-dropped"
  /**
   * Part of a footnote was dropped: a block a footnote cannot hold (a table),
   * a second definition claiming a number that is already taken, or a marker
   * that would have expanded a note into itself.
   */
  | "footnote-content-dropped"
  /** A footnote definition appeared somewhere other than the document root. */
  | "footnote-misplaced"
  /**
   * A `[^x]` reference had no definition to point at. The marker is still
   * drawn, but there is no note behind it — see {@link RenderOptions.footnotes}.
   */
  | "footnote-unresolved"
  /**
   * A `[^x]: …` definition that nothing leads to — neither a marker in the body
   * nor one inside another note. Word draws a footnote only where a reference
   * is, and drops an unreferenced one on the next save, so its text was not
   * written rather than written somewhere no reader would ever reach.
   */
  | "footnote-unreferenced"
  /**
   * A `TableOfContents` field was emitted. **A `TOC` field has no entries until
   * the reader updates it**; see {@link RenderOptions.toc}.
   */
  | "toc-needs-update"
  /** `options.highlighter` returned a Promise to the synchronous renderer. */
  | "highlighter-async"
  /** A {@link Highlighter} threw or rejected; that block renders plain monospace. */
  | "highlighter-failed"
  /**
   * A {@link Highlighter}'s spans did not reproduce the block's source, so the
   * source was restored. See {@link HighlightSpan} for the contract they broke.
   */
  | "highlighter-mismatch"
  /** A list nested deeper than OOXML's nine levels and was clamped to the ninth. */
  | "list-depth-truncated"
  /** An indent would have started past the right margin and was clamped to it. */
  | "indent-clamped"
  /** An image was taller than the text column and was scaled down to fit it. */
  | "image-oversized"
  /**
   * Right-to-left text was rendered into a left-to-right document. Set
   * {@link RenderOptions.direction} to `"rtl"`.
   */
  | "rtl-not-enabled"
  /**
   * A string reached a render sink carrying characters XML 1.0 cannot hold, and
   * each was replaced with `U+FFFD`. Unreachable through `parseMarkdown`, which
   * sanitizes first; a consumer-built model is what raises it. See
   * `src/xml-text.ts` for why the alternative is a file Word will not open.
   */
  | "unrepresentable-character"
  /** Two marks OOXML cannot combine met on one run; one of them was dropped. */
  | "conflicting-marks"
  /** A table had no cells at all and was skipped. */
  | "table-empty";

/**
 * Which render warnings mean content is missing from the `.docx`.
 *
 * Total by construction, and the only place the question is answered — see
 * `src/warnings.ts` for where the error/notice line is drawn.
 */
const RENDER_WARNING_SEVERITY: Readonly<Record<RenderWarningCode, WarningSeverity>> = {
  // Content is gone or stood in for.
  "image-unresolved": "error",
  "image-invalid": "error",
  "image-resolver-failed": "error",
  "footnote-content-dropped": "error",
  "footnote-misplaced": "error",
  "footnote-unresolved": "error",
  "footnote-unreferenced": "error",
  "table-empty": "error",
  // Everything the reader can still see, expressed differently.
  "html-block": "notice",
  "html-inline": "notice",
  "math-unconverted": "notice",
  "link-unresolved": "notice",
  "link-title-dropped": "notice",
  "code-meta-dropped": "notice",
  "highlighter-async": "notice",
  "highlighter-failed": "notice",
  "highlighter-mismatch": "notice",
  "list-depth-truncated": "notice",
  "indent-clamped": "notice",
  "image-oversized": "notice",
  "rtl-not-enabled": "notice",
  "unrepresentable-character": "notice",
  "conflicting-marks": "notice",
  "toc-needs-update": "notice",
};

/**
 * Every render warning code, so the set can be enumerated rather than guessed.
 *
 * Derived from the severity table, which the compiler already forces to be
 * total — a new code is a type error until it appears there, and then it
 * appears here too. A host building a "3 problems, 5 notes" summary or a
 * per-code label table needs the list; so does the test that checks severity is
 * answerable for all of them.
 */
export const RENDER_WARNING_CODES: readonly RenderWarningCode[] = Object.freeze(
  Object.keys(RENDER_WARNING_SEVERITY) as RenderWarningCode[],
);

/** The severity of a render warning code. Total by construction. */
export function renderWarningSeverity(code: RenderWarningCode): WarningSeverity {
  return RENDER_WARNING_SEVERITY[code];
}

/** A non-fatal problem encountered while rendering. See {@link BaseWarning}. */
export interface RenderWarning extends BaseWarning {
  readonly code: RenderWarningCode;
}

/** Called once per {@link RenderWarning}. Must not throw; if it does, it is ignored. */
export type RenderWarningHandler = (warning: RenderWarning) => void;

/** Builds a warning with the severity its code implies, and hands it over safely. */
export function reportRenderWarning(
  handler: RenderWarningHandler | null | undefined,
  code: RenderWarningCode,
  message: string,
): void {
  if (handler === null || handler === undefined) return;
  try {
    handler({ code, severity: renderWarningSeverity(code), message });
  } catch {
    // A host whose logger throws must not take the document down with it.
  }
}

/* -------------------------------------------------------------------------- */
/* Policies                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * What to do with `htmlBlock` / `htmlInline` nodes.
 *
 * The model keeps raw HTML rather than discarding it at parse time precisely so
 * that this decision is visible and configurable instead of silent.
 *
 * - `"raw"` (default) — emit the source verbatim in a monospace run, so the
 *   user can see what was not converted and fix it.
 * - `"drop"` — emit nothing. Still warns.
 */
export type HtmlPolicy = "raw" | "drop";

/**
 * What a `softBreak` (a newline inside a paragraph) becomes.
 *
 * - `"space"` (default) — CommonMark semantics: collapse to a single space.
 * - `"break"` — preserve the line break as `<w:br/>`. Matches what many people
 *   expect from hard-wrapped LLM output.
 * - `"ignore"` — emit nothing at all.
 */
export type SoftBreakPolicy = "space" | "break" | "ignore";

/**
 * The document's base writing direction.
 *
 * OOXML resolves a paragraph's *base direction* from `<w:bidi/>`, not from the
 * characters in it, and the Unicode Bidirectional Algorithm then lays the run
 * out relative to that base. Leaving it at `"ltr"` for Arabic or Hebrew prose
 * is not cosmetic: trailing punctuation and mixed Latin/Arabic segments resolve
 * on the wrong side of the line, list markers stay on the left, and the
 * paragraph aligns left.
 *
 * - `"ltr"` (default) — nothing is emitted, which is what every Latin-script
 *   document wants. Rendering right-to-left text under it raises a single
 *   `rtl-not-enabled` warning rather than silently mis-setting the page.
 * - `"rtl"` — prose paragraphs carry `<w:bidi/>`, their runs carry `<w:rtl/>`,
 *   and tables carry `<w:bidiVisual/>` so column order mirrors too. Code
 *   blocks, raw HTML and display math stay left-to-right, because their
 *   content is not prose.
 */
export type TextDirection = "ltr" | "rtl";

/**
 * What a `[^1]` reference and its `[^1]: …` definition become.
 *
 * - `true` (default) — real Word footnotes: a superscript, clickable
 *   `<w:footnoteReference>` in the body and the note's text in
 *   `word/footnotes.xml`. Word numbers, positions and renumbers them itself.
 * - `false` — the note is spliced into the sentence it was attached to, in
 *   parentheses. Nothing is dropped: a reader with no footnote pane (a
 *   converter, a plain-text extraction, a document destined for a slide) still
 *   sees every word the author wrote, just in a different place.
 */
export type FootnotePolicy = boolean;

/** How a page number is worded in the footer. */
export type PageNumberFormat =
  /** `7` */
  | "number"
  /** `Page 7` */
  | "page-x"
  /** `Page 7 of 12` */
  | "page-x-of-y";

/** Page numbers in the footer of every page. See {@link RenderOptions.pageNumbers}. */
export interface PageNumbersInit {
  /** Defaults to `"number"`. */
  readonly format?: PageNumberFormat | undefined;
  /** Defaults to `"center"`. */
  readonly alignment?: "left" | "center" | "right" | undefined;
}

/** A fully resolved {@link PageNumbersInit}. */
export interface PageNumberSettings {
  readonly format: PageNumberFormat;
  readonly alignment: "left" | "center" | "right";
}

/**
 * A native Word table-of-contents field. See {@link RenderOptions.toc}.
 *
 * **The field is empty until it is updated.** OOXML stores a `TOC` field as an
 * *instruction*, not as a list of entries; the entries are computed by the word
 * processor from the document's heading outline. downword marks the field
 * dirty and sets `w:updateFields`, which is everything the format allows a
 * generator to do — but the reader still has to act on it:
 *
 *  - **Word** offers "This document contains fields that may refer to other
 *    files. Update?" on open; answering yes fills the TOC in. Otherwise:
 *    right-click the field -> **Update Field**, or select it and press **F9**.
 *  - **LibreOffice Writer**: Tools -> Update -> Indexes and Tables.
 *  - **Google Docs / Pages / Quick Look / most converters**: they do not run
 *    fields at all, and will show the placeholder or nothing.
 */
export interface TocInit {
  /** Shallowest heading level listed, 1-6. Defaults to `1`. */
  readonly minLevel?: number | undefined;
  /** Deepest heading level listed, 1-6. Defaults to `3`. */
  readonly maxLevel?: number | undefined;
  /**
   * Heading printed above the field, or `null` for none. Defaults to
   * `"Contents"`.
   *
   * Styled `TOCHeading` — Word's own style id for exactly this, based on
   * `Heading1` but with `<w:outlineLvl w:val="9"/>` so the contents heading
   * does not list itself.
   */
  readonly title?: string | null | undefined;
}

/** A fully resolved {@link TocInit}. */
export interface TocSettings {
  readonly minLevel: number;
  readonly maxLevel: number;
  readonly title: string | null;
}

/** A named paper size, or explicit portrait dimensions in twips. */
export type PageSize = "A4" | "Letter" | { readonly width: number; readonly height: number };

/** Page geometry. All measurements are twips. */
export interface PageInit {
  /** Defaults to `"A4"` (11906 x 16838 twips), matching docx's own default. */
  readonly size?: PageSize;
  readonly orientation?: "portrait" | "landscape";
  /** Defaults to 1 inch (1440) on every side. */
  readonly margin?: {
    readonly top?: number;
    readonly right?: number;
    readonly bottom?: number;
    readonly left?: number;
  };
}

/** A fully resolved {@link PageInit}. */
export interface PageSettings {
  /** Portrait width in twips (docx swaps w/h itself when landscape). */
  readonly width: number;
  readonly height: number;
  readonly orientation: "portrait" | "landscape";
  readonly margin: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  /** Usable text width: the long or short edge, less both side margins. */
  readonly contentWidth: number;
  /**
   * Usable text height: the other edge, less the top and bottom margins.
   *
   * Word does not split an inline picture across a page boundary and does not
   * shrink one to fit, so anything below the first boundary simply is not
   * drawn. This is the cap the renderer scales images against; see
   * `render/inline.ts`.
   */
  readonly contentHeight: number;
}

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

/** Options for {@link import("./index.js").renderDocument}. */
export interface RenderOptions {
  /** Design tokens, merged over `DEFAULT_THEME`. */
  readonly theme?: ThemeInit | undefined;
  /** Page size and margins. */
  readonly page?: PageInit | undefined;
  /**
   * Pre-computed highlight spans from {@link prepareHighlights}. Preferred over
   * {@link RenderOptions.highlighter} because it cannot be async-surprised.
   */
  readonly highlights?: HighlightMap | undefined;
  /**
   * A highlighter called inline. Only a *synchronous* return value can be used;
   * a Promise raises a `highlighter-async` warning and the block falls back to
   * plain monospace. Use {@link prepareHighlights} for async highlighters.
   */
  readonly highlighter?: Highlighter | undefined;
  /**
   * Pre-resolved images from {@link prepareImages}. Consulted when an
   * {@link ImageNode} does not already carry `resolved` bytes.
   */
  readonly images?: ImageMap | undefined;
  /** Raw-HTML policy. Defaults to `"raw"`. */
  readonly html?: HtmlPolicy | undefined;
  /** Soft-break policy. Defaults to `"space"`. */
  readonly softBreak?: SoftBreakPolicy | undefined;
  /** Base writing direction. Defaults to `"ltr"`. See {@link TextDirection}. */
  readonly direction?: TextDirection | undefined;
  /**
   * Spaces a tab expands to inside a code block. Defaults to `4`.
   *
   * Must be an integer in `[0, 64]`; `0` leaves tabs alone. Anything else
   * raises a `DownwordError` with code `"invalid-options"` — a tab expanded to
   * a billion spaces is not a document, it is an out-of-memory crash.
   */
  readonly tabSize?: number | undefined;
  /**
   * When the document's metadata carries a title, emit it as a `Title`-styled
   * paragraph at the top of the body. Defaults to `true`.
   */
  readonly titleBlock?: boolean | undefined;
  /**
   * Real Word footnotes, or parenthetical text. Defaults to `true`. See
   * {@link FootnotePolicy}.
   */
  readonly footnotes?: FootnotePolicy | undefined;
  /**
   * Emit a native Word table-of-contents field above the body. Defaults to
   * `false`.
   *
   * `true` means `{ minLevel: 1, maxLevel: 3, title: "Contents" }`.
   *
   * **Read {@link TocInit} before enabling this.** The field is empty until the
   * reader updates it (Word: right-click -> Update Field, or answer yes to the
   * prompt on open). A `toc-needs-update` notice is raised for every document
   * that gets one, so a host can say so in its own UI.
   */
  readonly toc?: boolean | TocInit | undefined;
  /**
   * Put a page number in the footer of every page. Defaults to `false`.
   *
   * `true` means `{ format: "number", alignment: "center" }`. Unlike
   * {@link RenderOptions.toc}, `PAGE`/`NUMPAGES` fields are computed during
   * layout by every reader that paginates at all, so these need no update step.
   */
  readonly pageNumbers?: boolean | PageNumbersInit | undefined;
  /**
   * OOXML `dc:subject` — the **Subject** field in Word's properties panel.
   *
   * Every other core property is lifted from `doc.metadata`, but the model has
   * no `subject` field: markdown has nothing to lift one *from*, and
   * frontmatter effectively never carries it. Exposing it as an option lets a
   * host that does know the subject write the genuine core property rather
   * than a look-alike custom property. Defaults to unset.
   */
  readonly subject?: string | undefined;
  /** Called once per non-fatal problem. A handler that throws is ignored. */
  readonly onWarning?: RenderWarningHandler | undefined;
}

/**
 * Options shared by the render stage's two async preludes.
 *
 * Both {@link import("./prepare.js").prepareHighlights} and
 * {@link import("./prepare.js").prepareImages} degrade one node at a time
 * rather than failing the document, which is only defensible if the caller can
 * find out that they did.
 */
export interface PrepareOptions {
  /** Called once per node whose enhancement was lost. Never throws through. */
  readonly onWarning?: RenderWarningHandler | undefined;
}

/** {@link RenderOptions} with every default filled in. */
export interface ResolvedRenderOptions {
  readonly html: HtmlPolicy;
  readonly softBreak: SoftBreakPolicy;
  readonly direction: TextDirection;
  readonly tabSize: number;
  readonly titleBlock: boolean;
  readonly footnotes: FootnotePolicy;
  /** `null` when no table of contents was asked for. */
  readonly toc: TocSettings | null;
  /** `null` when the footer carries no page number. */
  readonly pageNumbers: PageNumberSettings | null;
}
