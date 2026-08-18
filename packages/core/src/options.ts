/**
 * `convert()`'s options, and the mapping from them onto the four subsystems.
 *
 * downword is four independent pieces — a parser, an image pass, a highlighting
 * pass and a renderer — each with its own, deliberately wide, option type.
 * {@link ConvertOptions} is the narrow front door: one flat, documented object
 * that covers what a person converting a markdown file actually wants to
 * decide, and {@link resolveConvertOptions} is the single place that translates
 * it. Anything the front door does not cover is still reachable by driving
 * `parseMarkdown` and `renderDocument` yourself.
 *
 * Two rules this file keeps:
 *
 *  - **Every option has a documented default**, and the default is stated in
 *    the TSDoc rather than left to be inferred from the code.
 *  - **Bad values fail loudly.** An out-of-range margin or an unknown theme
 *    name raises a {@link DownwordError} with code `"invalid-options"` instead
 *    of producing a document Word quietly refuses to open.
 */

import { metadata as buildMetadata, type DocumentMetadata } from "./model.js";
import { invalid, oneOf, twips } from "./validate.js";
import { PRINT_CODE_PALETTE } from "./highlight/palette.js";
import type { MarkdownItPlugin, ParseOptions, ParseWarning } from "./parse/index.js";
import type {
  Highlighter,
  HtmlPolicy,
  ImageResolver,
  PageInit,
  PageNumbersInit,
  PageSize,
  RenderOptions,
  RenderWarning,
  SoftBreakPolicy,
  TextDirection,
  Theme,
  ThemeInit,
  TocInit,
} from "./render/index.js";
import { DEFAULT_THEME, resolveTheme } from "./render/index.js";
// The presets themselves, straight from the module that owns them: they are
// theme *values*, and `render/index.ts` re-exports the theme machinery rather
// than every table built with it.
import { ACADEMIC_DOUBLE_THEME, ACADEMIC_THEME, GITHUB_THEME } from "./render/theme.js";
import type { ImageDiagnostic } from "./images/index.js";
import { sanitizeXmlText } from "./xml-text.js";

/* -------------------------------------------------------------------------- */
/* Themes                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * A built-in theme.
 *
 * Three of the five are *looks* — the same document, dressed differently — and
 * two are variants of one of those. All five differ **only in the generated
 * `word/styles.xml`**, with the two documented exceptions listed on
 * `GITHUB_THEME` in `render/theme.ts`; `tests/themes.test.ts` renders one
 * fixture under each and asserts the `word/document.xml` it produces is the
 * same file. That is what makes a converted document restylable in Word rather
 * than merely pre-styled.
 *
 * - `"default"` — Word's own look: Aptos with a Calibri fallback, Word's
 *   heading blues, 1.15 leading. A GitHub-flavoured palette covers the things
 *   Word has no opinion about (code, tables, quotes); its code inks clear WCAG
 *   **AA** (4.5:1) against the code block's own shading, floor 4.74:1.
 * - `"github"` — a README as Word can render one: Segoe UI (fallback Arial),
 *   near-black headings rather than tinted ones, single leading and smaller
 *   gaps throughout.
 * - `"academic"` — Times New Roman 12 pt (fallback Cambria), black headings,
 *   1.5 leading, Courier New for code.
 * - `"academic-double"` — the same, double-spaced. One token (`spacing.line`,
 *   480), because that is what a thesis office asks for and nothing else about
 *   the manuscript should move with it.
 * - `"print"` — `"default"` with the syntax-highlighting palette **replaced** by
 *   one measured against that shading rather than against white: every ink
 *   clears WCAG **AAA** (7:1), the floor being 7.13:1. Scopes it does not name
 *   fall through to the code block's own near-black (13.76:1) rather than
 *   inheriting a screen ink. Use it for documents that get printed or
 *   photocopied.
 *
 * The palette is a separate axis from the look, and composes with it:
 * `{ ...THEMES.academic, codePalette: THEMES.print.codePalette }`. Whichever
 * you pick governs a `highlighter`'s output too, so
 * `convert(md, { theme, highlighter: createHighlighter() })` colours real code
 * blocks with the theme you asked for. Weight and slant (bold keywords, italic
 * comments) come from the highlighter under every theme, because a `.docx` is a
 * print artefact whichever palette it uses.
 *
 * Pass a {@link ThemeInit} instead to override individual tokens; pass a
 * resolved {@link Theme} (including one of {@link THEMES}) to start from it.
 */
export type ThemeName = "default" | "github" | "academic" | "academic-double" | "print";

/** The built-in themes, resolved. See {@link ThemeName}. */
export const THEMES: Readonly<Record<ThemeName, Theme>> = Object.freeze({
  default: DEFAULT_THEME,
  github: GITHUB_THEME,
  academic: ACADEMIC_THEME,
  "academic-double": ACADEMIC_DOUBLE_THEME,
  // `codePalette` replaces rather than merges (see ThemeInit), which is what
  // keeps the AAA floor above true: a merge left eight screen inks in place,
  // the darkest of which (116329, 6.94:1) missed it.
  print: resolveTheme({ codePalette: PRINT_CODE_PALETTE }),
});

/* -------------------------------------------------------------------------- */
/* Warnings                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Which stage of the pipeline raised a {@link ConvertWarning}.
 *
 * Also the discriminant of the union, because the useful *detail* differs per
 * stage: a parse warning knows a source line, an image diagnostic knows a URL
 * and whether the picture survived, a render warning knows neither.
 */
export type ConvertWarningStage = "parse" | "image" | "render";

/**
 * A non-fatal problem found while converting.
 *
 * Warnings never stop a conversion — the document is always produced. They
 * exist so a host can tell the user *"three images could not be loaded"*
 * instead of silently handing back a file with holes in it.
 */
export type ConvertWarning =
  | ({ readonly stage: "parse" } & ParseWarning)
  | ({ readonly stage: "image" } & ImageDiagnostic)
  | ({ readonly stage: "render" } & RenderWarning);

/** Called once per {@link ConvertWarning}. Must not throw; if it does, it is ignored. */
export type ConvertWarningHandler = (warning: ConvertWarning) => void;

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

/** Page margins in twips, or one number applied to all four sides. */
export type MarginsInit =
  | number
  | {
      readonly top?: number | undefined;
      readonly right?: number | undefined;
      readonly bottom?: number | undefined;
      readonly left?: number | undefined;
    };

/**
 * Core document properties, as Word shows them under File → Info.
 *
 * Anything set here is written into `docProps/core.xml`; nothing is inferred
 * from the markdown itself (downword does not parse frontmatter — that is a
 * host concern, and YAML is not a dependency of this package).
 */
export interface ConvertMetadata {
  /** `dc:title`. Shown as **Title**. Defaults to unset. */
  readonly title?: string | undefined;
  /** `dc:creator`. Shown as **Author**. Defaults to unset. */
  readonly author?: string | undefined;
  /** `dc:subject`. Shown as **Subject**. Defaults to unset. */
  readonly subject?: string | undefined;
  /** `dc:description`. Shown as **Comments**. Defaults to unset. */
  readonly description?: string | undefined;
  /** `cp:keywords`, joined with `", "`. Defaults to none. */
  readonly keywords?: readonly string[] | undefined;
  /**
   * Extra properties, written as OOXML custom document properties (Word shows
   * them under File → Info → Properties → Advanced). Defaults to none.
   */
  readonly custom?: Readonly<Record<string, string>> | undefined;
}

/**
 * How raw HTML embedded in the markdown is treated.
 *
 * LLM output is full of stray `<br>`, `<sub>` and `<div align="center">`, and
 * OOXML has no way to embed HTML, so there is no lossless answer — only three
 * honest ones.
 *
 * - `"escape"` (default) — markdown-it's own safe default: the markup is not
 *   parsed, so it survives as literal text and the reader sees `<br>`.
 * - `"keep"` — the markup is parsed into HTML nodes and re-emitted verbatim in
 *   a monospace run, so it is visibly *unconverted* rather than pretending.
 * - `"drop"` — the markup is parsed and then removed. The cleanest-looking
 *   result, and the one to pick for machine-generated markdown.
 *
 * `"keep"` and `"drop"` raise a warning per construct — `raw-html` from the
 * parser and `html-block`/`html-inline` from the renderer — so nothing
 * disappears silently. `"escape"` raises none, and needs none: markdown-it
 * never produces an HTML token, the angle brackets are simply text, and there
 * is nothing for the reader to have lost.
 */
export type HtmlHandling = "escape" | "keep" | "drop";

/**
 * What a single newline inside a paragraph becomes.
 *
 * - `"collapse"` (default) — CommonMark semantics: a space. Correct for
 *   hard-wrapped prose.
 * - `"preserve"` — a real line break, which is what people expect when the
 *   markdown was wrapped for a chat window rather than for a paragraph.
 */
export type LineBreakHandling = "collapse" | "preserve";

/**
 * Options for {@link import("./convert.js").convert}.
 *
 * Every field is optional and every default is documented below. Nothing here
 * performs I/O unless you ask it to: no image is fetched over the network
 * unless {@link ConvertOptions.allowRemoteImages} is set, and highlight.js is
 * never loaded unless you pass a {@link ConvertOptions.highlighter}.
 */
export interface ConvertOptions {
  /* --- appearance ------------------------------------------------------- */

  /**
   * Design tokens: fonts, colours, sizes, spacing, list glyphs.
   *
   * A {@link ThemeName} selects a built-in; a {@link ThemeInit} is merged one
   * level deep over the `"default"` theme. Defaults to `"default"`.
   *
   * ```ts
   * import { convert, THEMES } from "@ksprtech/downword";
   * await convert("# hi", { theme: { fonts: { body: "Georgia" } } });
   * await convert("# hi", { theme: { ...THEMES.print, sizes: { ...THEMES.print.sizes, body: 24 } } });
   * ```
   */
  readonly theme?: ThemeName | ThemeInit | undefined;
  /**
   * Paper size: `"A4"` (default), `"Letter"`, or explicit dimensions in twips.
   *
   * Twips are twentieths of a point: 1440 = 1 inch. A4 is 11906 x 16838.
   */
  readonly pageSize?: PageSize | undefined;
  /** Page orientation. Defaults to `"portrait"`. */
  readonly orientation?: "portrait" | "landscape" | undefined;
  /**
   * Page margins in twips, or one number for all four sides. Defaults to
   * `1440` (1 inch), which is Word's own default.
   *
   * ```ts
   * import { convert, inchesToTwips } from "@ksprtech/downword";
   * await convert("# hi", { margins: inchesToTwips(0.75) });
   * await convert("# hi", { margins: { left: 2160, right: 2160 } });
   * ```
   */
  readonly margins?: MarginsInit | undefined;
  /**
   * The document's base writing direction. Defaults to `"ltr"`.
   *
   * `"rtl"` is what an Arabic, Hebrew, Persian or Urdu document needs: prose
   * paragraphs get `<w:bidi/>`, their runs get `<w:rtl/>` (which also makes
   * Word use the complex-script face and size rather than the Latin ones), and
   * tables get `<w:bidiVisual/>` so the column order mirrors too. Code blocks,
   * raw HTML and display math stay left-to-right, because their content is not
   * prose.
   *
   * Leaving it at `"ltr"` for right-to-left prose is not merely cosmetic — the
   * Unicode Bidirectional Algorithm resolves trailing punctuation and mixed
   * Latin/Arabic segments against the *paragraph's* base direction — so doing
   * it raises one `rtl-not-enabled` notice rather than passing in silence.
   */
  readonly direction?: TextDirection | undefined;
  /**
   * Spaces a tab inside a fenced code block expands to. Defaults to `4`.
   *
   * A literal tab in `<w:t>` is not a tab stop, and `<w:tab/>` measures against
   * the paragraph's tab stops rather than the listing's column grid, so
   * expanding is the only way indentation survives in monospace. `0` leaves
   * tabs alone. Must be an integer in `[0, 64]`.
   */
  readonly tabSize?: number | undefined;
  /**
   * Repeat {@link ConvertMetadata.title} as a `Title`-styled paragraph at the
   * top of the document. Defaults to `false`.
   *
   * (`renderDocument`'s own default is `true`. It is flipped here because a
   * title set through `convert()` is normally there for the *file's*
   * properties, and the markdown usually already opens with its own `#`
   * heading — which would then appear twice.)
   */
  readonly titleBlock?: boolean | undefined;
  /**
   * Put a native Word table-of-contents field at the top. Defaults to `false`.
   *
   * `true` means `{ minLevel: 1, maxLevel: 3, title: "Contents" }`; pass a
   * {@link TocInit} to change the range or the heading.
   *
   * **The field is empty until the reader updates it.** OOXML stores a `TOC`
   * field as an instruction, not as a list of entries, and the entries are
   * computed by the word processor from the document's heading outline.
   * downword marks the field dirty and sets `<w:updateFields/>`, which is
   * everything the format allows a generator to do — Word then offers to update
   * on open (or right-click the field -> **Update Field**, or **F9**);
   * LibreOffice needs Tools -> Update -> Indexes and Tables; Google Docs, Pages
   * and most converters do not run fields at all and will show nothing.
   *
   * Every document that gets one raises a `toc-needs-update` notice, so a host
   * can say this in its own words.
   */
  readonly toc?: boolean | TocInit | undefined;
  /**
   * Put a page number in the footer of every page. Defaults to `false`.
   *
   * `true` means `{ format: "number", alignment: "center" }`. Unlike
   * {@link ConvertOptions.toc}, `PAGE`/`NUMPAGES` are computed during layout by
   * every reader that paginates, so these need no update step and raise no
   * warning.
   */
  readonly pageNumbers?: boolean | PageNumbersInit | undefined;

  /* --- content ---------------------------------------------------------- */

  /** Core document properties. See {@link ConvertMetadata}. Defaults to none. */
  readonly metadata?: ConvertMetadata | undefined;
  /**
   * What `[^1]` footnotes become. Defaults to `true`.
   *
   * - `true` — real Word footnotes: a superscript, clickable reference in the
   *   body and the note's text in `word/footnotes.xml`, which Word numbers,
   *   positions and renumbers itself.
   * - `"inline"` — the syntax is still parsed, but each note is spliced into
   *   the sentence that cited it, in parentheses. For a document destined for
   *   a reader with no footnote pane: a converter, a plain-text extraction, a
   *   slide.
   * - `false` — the syntax is not parsed at all and `[^1]` stays literal text.
   */
  readonly footnotes?: boolean | "inline" | undefined;
  /** Raw-HTML handling. Defaults to `"escape"`. See {@link HtmlHandling}. */
  readonly html?: HtmlHandling | undefined;
  /**
   * Turn bare URLs and e-mail addresses into hyperlinks. Defaults to `true`,
   * because pasted LLM output is full of unlinked URLs.
   */
  readonly linkify?: boolean | undefined;
  /** Smart quotes, en/em dashes and ellipses. Defaults to `false`. */
  readonly typographer?: boolean | undefined;
  /** Single-newline handling. Defaults to `"collapse"`. See {@link LineBreakHandling}. */
  readonly lineBreaks?: LineBreakHandling | undefined;
  /**
   * Extra markdown-it plugins, applied in order after the built-ins. Defaults
   * to none.
   *
   * This is the extension point for syntax downword does not know: a plugin
   * emitting `math_inline`/`math_block` tokens becomes math nodes, for example.
   * Plugins that take options are passed as a closure, which keeps the type
   * free of `any`:
   *
   * ```ts
   * import { convert, type MarkdownItPlugin } from "@ksprtech/downword";
   *
   * const deflist: MarkdownItPlugin = (md) => {
   *   md.block.ruler.before("paragraph", "noop", () => false);
   * };
   * await convert("term\n: definition", { plugins: [deflist] });
   * ```
   */
  readonly plugins?: readonly MarkdownItPlugin[] | undefined;

  /* --- extensions ------------------------------------------------------- */

  /**
   * Colours fenced code blocks. Defaults to none — code renders as plain
   * monospace and highlight.js is never loaded.
   *
   * ```ts
   * import { convert } from "@ksprtech/downword";
   * import { createHighlighter } from "@ksprtech/downword/highlight";
   *
   * await convert("```ts\nconst x = 1;\n```", { highlighter: createHighlighter() });
   * ```
   */
  readonly highlighter?: Highlighter | undefined;
  /**
   * Turns `![](src)` destinations into embeddable bytes.
   *
   * Defaults to `createImageResolver({ allowRemote: allowRemoteImages })`,
   * which reads `data:` and `blob:` URLs only. Pass
   * `createNodeImageResolver()` from `@ksprtech/downword/images/node` to read local
   * files, `NULL_IMAGE_RESOLVER` to resolve nothing at all, or your own.
   *
   * Whatever you pass, an image that fails to resolve costs exactly one
   * placeholder and one warning — never the document.
   */
  readonly imageResolver?: ImageResolver | undefined;
  /**
   * Let the default resolver fetch `http(s)` images. Defaults to **`false`**:
   * out of the box a conversion cannot make a network request.
   *
   * Ignored when {@link ConvertOptions.imageResolver} is supplied — that
   * resolver's own policy applies instead.
   */
  readonly allowRemoteImages?: boolean | undefined;

  /* --- diagnostics ------------------------------------------------------ */

  /**
   * Called once per non-fatal problem, from every stage. Defaults to none
   * (problems are silent). See {@link ConvertWarning}.
   *
   * A handler that throws is ignored; it cannot fail the conversion.
   */
  readonly onWarning?: ConvertWarningHandler | undefined;
}

/* -------------------------------------------------------------------------- */
/* Resolution                                                                  */
/* -------------------------------------------------------------------------- */

/** {@link ConvertOptions}, translated into what each subsystem takes. */
export interface ResolvedConvertOptions {
  readonly parse: ParseOptions;
  readonly render: RenderOptions;
  /** `null` when the caller supplied no highlighter: code renders plain. */
  readonly highlighter: Highlighter | null;
  /** The caller's resolver, or `null` to build the default one lazily. */
  readonly imageResolver: ImageResolver | null;
  /** Whether the default resolver may fetch `http(s)`. */
  readonly allowRemoteImages: boolean;
}

/** Word's default page margin: one inch. */
const DEFAULT_MARGIN_TWIPS = 1440;

const HTML_POLICIES: Readonly<Record<HtmlHandling, { parse: boolean; render: HtmlPolicy }>> = {
  escape: { parse: false, render: "raw" },
  keep: { parse: true, render: "raw" },
  drop: { parse: true, render: "drop" },
};

const SOFT_BREAKS: Readonly<Record<LineBreakHandling, SoftBreakPolicy>> = {
  collapse: "space",
  preserve: "break",
};

/**
 * Splits {@link ConvertOptions.footnotes} across the two stages that answer it.
 *
 * One option, two decisions, and they are not the same decision: the *parser*
 * decides whether `[^1]` is syntax at all, and the *renderer* decides what the
 * resulting node becomes. `"inline"` is the combination the two-boolean shape
 * could not express — parse the syntax, then splice each note into the sentence
 * that cited it instead of putting it in `word/footnotes.xml`.
 *
 * `false` turns the parser off, and the renderer's own policy is then moot:
 * there are no footnote nodes for it to have an opinion about.
 */
function resolveFootnotes(value: boolean | "inline" | undefined): {
  readonly parse: boolean;
  readonly render: boolean;
} {
  if (value === undefined || value === true) return { parse: true, render: true };
  if (value === false) return { parse: false, render: true };
  if (value === "inline") return { parse: true, render: false };
  throw invalid(`options.footnotes must be true, false or "inline", got ${JSON.stringify(value)}`);
}

const ORIENTATIONS = ["portrait", "landscape"] as const;

function resolveMargins(init: MarginsInit | undefined): PageInit["margin"] {
  if (init === undefined) return undefined;
  if (typeof init === "number") {
    const all = twips(init, "margins", DEFAULT_MARGIN_TWIPS);
    return { top: all, right: all, bottom: all, left: all };
  }
  if (typeof init !== "object" || init === null) {
    throw invalid(`options.margins must be a number or an object, got ${JSON.stringify(init)}`);
  }
  return {
    top: twips(init.top, "margins.top", DEFAULT_MARGIN_TWIPS),
    right: twips(init.right, "margins.right", DEFAULT_MARGIN_TWIPS),
    bottom: twips(init.bottom, "margins.bottom", DEFAULT_MARGIN_TWIPS),
    left: twips(init.left, "margins.left", DEFAULT_MARGIN_TWIPS),
  };
}

function resolvePageSize(size: PageSize | undefined): PageSize {
  if (size === undefined) return "A4";
  if (typeof size === "string") return oneOf(size, ["A4", "Letter"] as const, "pageSize", "A4");
  if (typeof size !== "object" || size === null) {
    throw invalid(`options.pageSize must be "A4", "Letter" or { width, height }`);
  }
  const width = twips(size.width, "pageSize.width", 0);
  const height = twips(size.height, "pageSize.height", 0);
  if (width <= 0 || height <= 0) {
    throw invalid(`options.pageSize dimensions must be positive twips, got ${width}x${height}`);
  }
  return { width, height };
}

function resolveThemeOption(theme: ThemeName | ThemeInit | undefined): ThemeInit {
  if (theme === undefined) return THEMES.default;
  if (typeof theme === "string") {
    if (!Object.prototype.hasOwnProperty.call(THEMES, theme)) {
      throw invalid(
        `options.theme must be one of ${Object.keys(THEMES)
          .map((name) => `"${name}"`)
          .join(", ")}, or a ThemeInit object; got ${JSON.stringify(theme)}`,
      );
    }
    return THEMES[theme];
  }
  if (typeof theme !== "object" || theme === null) {
    throw invalid(`options.theme must be a theme name or a ThemeInit object`);
  }
  return theme;
}

/**
 * Metadata reaches `docProps/*.xml` without passing through the parser, so it
 * gets the same XML-safety sweep the markdown does. See `src/xml-text.ts`; a
 * single control character in a pasted title is enough to make Word reject the
 * whole file.
 */
function safe(value: string): string {
  return sanitizeXmlText(value);
}

function resolveMetadata(init: ConvertMetadata | undefined): DocumentMetadata {
  if (init === undefined) return buildMetadata();
  if (
    init.keywords !== undefined &&
    (!Array.isArray(init.keywords) || init.keywords.some((word) => typeof word !== "string"))
  ) {
    // Caught here rather than in the renderer, where it would surface as an
    // opaque `render-failed` from a `.join()` on something that is not a list.
    throw invalid(`options.metadata.keywords must be an array of strings`);
  }
  const custom: Record<string, string> = {};
  for (const [key, value] of Object.entries(init.custom ?? {})) {
    custom[safe(key)] = safe(String(value));
  }
  return buildMetadata({
    ...(init.title !== undefined ? { title: safe(init.title) } : {}),
    ...(init.author !== undefined ? { author: safe(init.author) } : {}),
    ...(init.description !== undefined ? { description: safe(init.description) } : {}),
    ...(init.keywords !== undefined ? { keywords: init.keywords.map(safe) } : {}),
    custom,
  });
}

/**
 * Translates {@link ConvertOptions} into the four subsystems' own options.
 *
 * Pure and synchronous: it validates, it does not convert. `emit` is the fan-in
 * every stage's warnings are funnelled through, wrapped so a handler that
 * throws cannot fail the conversion.
 *
 * @param options - The caller's options.
 * @param emit - Receives every {@link ConvertWarning}.
 * @throws {DownwordError} with code `"invalid-options"` for an out-of-range value.
 */
export function resolveConvertOptions(
  options: ConvertOptions,
  emit: ConvertWarningHandler,
): ResolvedConvertOptions {
  if (typeof options !== "object" || options === null) {
    throw invalid(`options must be an object, got ${JSON.stringify(options)}`);
  }

  const html = HTML_POLICIES[oneOf(options.html, ["escape", "keep", "drop"], "html", "escape")];
  const softBreak =
    SOFT_BREAKS[oneOf(options.lineBreaks, ["collapse", "preserve"], "lineBreaks", "collapse")];
  const orientation = oneOf(options.orientation, ORIENTATIONS, "orientation", "portrait");
  const margin = resolveMargins(options.margins);
  const rawSubject = options.metadata?.subject;
  const subject = rawSubject === undefined ? undefined : safe(rawSubject);
  const footnotes = resolveFootnotes(options.footnotes);

  const page: PageInit = {
    size: resolvePageSize(options.pageSize),
    orientation,
    ...(margin === undefined ? {} : { margin }),
  };

  const parse: ParseOptions = {
    html: html.parse,
    linkify: options.linkify ?? true,
    typographer: options.typographer ?? false,
    footnotes: footnotes.parse,
    metadata: resolveMetadata(options.metadata),
    ...(options.plugins === undefined ? {} : { plugins: options.plugins }),
    onWarning: (warning: ParseWarning) => {
      // `stage` last: it is this layer's fact, not the parser's to overwrite.
      emit({ ...warning, stage: "parse" });
    },
  };

  const render: RenderOptions = {
    theme: resolveThemeOption(options.theme),
    page,
    html: html.render,
    softBreak,
    // Left to `resolveOptions` in the renderer to validate, so that the two
    // public entry points cannot disagree about what a legal value is.
    ...(options.direction === undefined ? {} : { direction: options.direction }),
    ...(options.tabSize === undefined ? {} : { tabSize: options.tabSize }),
    titleBlock: options.titleBlock ?? false,
    footnotes: footnotes.render,
    // Both left to `resolveOptions` in the renderer to validate, for the same
    // reason `direction` is: one definition of a legal value, not two.
    ...(options.toc === undefined ? {} : { toc: options.toc }),
    ...(options.pageNumbers === undefined ? {} : { pageNumbers: options.pageNumbers }),
    ...(subject === undefined ? {} : { subject }),
    onWarning: (warning: RenderWarning) => {
      emit({ ...warning, stage: "render" });
    },
  };

  return {
    parse,
    render,
    highlighter: options.highlighter ?? null,
    imageResolver: options.imageResolver ?? null,
    allowRemoteImages: options.allowRemoteImages ?? false,
  };
}
