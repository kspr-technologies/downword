/**
 * The renderer's design tokens.
 *
 * Every number, colour, font and glyph the renderer can emit lives here, and
 * nothing else in `src/render/**` hard-codes one. The theme feeds two things:
 *
 *  - `buildStyles()`, which turns it into a real `word/styles.xml`; and
 *  - `createNumberingRegistry()`, which turns it into a real
 *    `word/numbering.xml`.
 *
 * That indirection is what makes the "no inline-only formatting" rule
 * enforceable: if a visual decision is expressible as a style property it is
 * declared here and emitted once into `styles.xml`, rather than being stamped
 * onto every paragraph.
 *
 * Sizes are in **half-points** (`w:sz`), spacing/indents in **twips**, border
 * weights in **eighth-points**, and colours are `RRGGBB` **without** a leading
 * `#` (docx rejects the hash).
 */

import type { HeadingLevel } from "../model.js";
import type { HighlightColor, LevelFormat } from "docx";

/* -------------------------------------------------------------------------- */
/* Leaf token types                                                            */
/* -------------------------------------------------------------------------- */

/** One of the 17 highlight colours OOXML allows for `<w:highlight>`. */
export type HighlightColorValue = (typeof HighlightColor)[keyof typeof HighlightColor];

/** One of the ~70 `<w:numFmt>` values OOXML allows for a numbering level. */
export type LevelFormatValue = (typeof LevelFormat)[keyof typeof LevelFormat];

/**
 * A per-script font chain: one face per OOXML script slot.
 *
 * OOXML has no font *stack* — `w:rFonts` names one face per script, and a
 * reader that cannot resolve it falls back to whatever its own substitution
 * table says. What the element does give is four slots, and naming a second,
 * more widely installed face in `cs` is the closest thing to a fallback the
 * format has: a reader without Aptos still gets Calibri for complex-script
 * text rather than a default nobody chose, and the pair documents the
 * designer's intent in the file itself.
 *
 * `eastAsia` is deliberately **not** part of this type. Leaving the slot unset
 * is what keeps CJK text on a CJK-capable face; pinning it to a Latin face (as
 * a bare `string` font does, because `docx` copies the name into all four
 * slots) is how CJK ends up rendered in a font that has no glyphs for it.
 */
export interface ThemeFontStack {
  /** Latin text. `w:ascii`. */
  readonly ascii: string;
  /** High-ANSI text. `w:hAnsi`. Normally the same face as {@link ThemeFontStack.ascii}. */
  readonly hAnsi: string;
  /** Complex-script text (Arabic, Hebrew, Thai), and the documented fallback face. `w:cs`. */
  readonly cs: string;
}

/**
 * One face, or a {@link ThemeFontStack}.
 *
 * A bare string is the ergonomic form and what `docx` expands into all four
 * script slots; a stack is what a theme uses to name a preferred face and a
 * fallback.
 */
export type ThemeFont = string | ThemeFontStack;

/**
 * Builds a two-face {@link ThemeFontStack}.
 *
 * @param preferred - The face to use for Latin text.
 * @param fallback - The more widely installed face, written to `w:cs`.
 */
export function fontStack(preferred: string, fallback: string): ThemeFontStack {
  return { ascii: preferred, hAnsi: preferred, cs: fallback };
}

/** Typefaces used across the document. */
export interface ThemeFonts {
  /** Body text; becomes `docDefaults`' `w:rFonts`. */
  readonly body: ThemeFont;
  /** Headings and the title. */
  readonly heading: ThemeFont;
  /** Preferred monospace face for code (`w:ascii`/`w:hAnsi`). */
  readonly mono: string;
  /**
   * Monospace face written to `w:cs`, which is the closest thing OOXML has to
   * a fallback chain: readers that cannot resolve {@link ThemeFonts.mono} fall
   * back to it rather than to a proportional font.
   */
  readonly monoFallback: string;
  /** Face used for task-list checkbox glyphs. */
  readonly symbol: string;
}

/** Colours, `RRGGBB`, no `#`. */
export interface ThemeColors {
  /** Headings 1, 2, 4 and 5. */
  readonly heading: string;
  /** Headings 3 and 6 (Word's own style set darkens the deeper levels). */
  readonly headingDeep: string;
  /** Hyperlink text. */
  readonly link: string;
  /** De-emphasised text: image placeholders, raw HTML. */
  readonly muted: string;
  /** Body text inside a block quote. */
  readonly quoteText: string;
  /** The vertical rule down the left of a block quote. */
  readonly quoteBorder: string;
  /** Text inside a fenced code block. */
  readonly codeText: string;
  /** Fenced code block background. */
  readonly codeBackground: string;
  /** Inline `code` text. */
  readonly inlineCodeText: string;
  /** Inline `code` background. */
  readonly inlineCodeBackground: string;
  /** Thematic break (`---`) rule. */
  readonly rule: string;
  /** Table grid lines. */
  readonly tableBorder: string;
  /** Header-row cell shading. */
  readonly tableHeaderBackground: string;
}

/** Font sizes, in half-points (22 = 11 pt). */
export interface ThemeSizes {
  readonly body: number;
  readonly title: number;
  readonly code: number;
  readonly footnote: number;
  /** Size of the paragraph mark on a thematic break; keeps the rule compact. */
  readonly rule: number;
}

/** Spacing and indentation, in twips (1440 = 1 inch), plus border weights. */
export interface ThemeSpacing {
  /** `w:spacing w:after` on a body paragraph. */
  readonly paragraphAfter: number;
  /** `w:spacing w:line` with `w:lineRule="auto"`; 240 = single. */
  readonly line: number;
  /** Indent added per list level, and the width of the number/bullet gutter. */
  readonly listIndent: number;
  /** `w:ind w:hanging` on a list paragraph. */
  readonly listHanging: number;
  /** `w:spacing w:after` between the items of a *loose* list. */
  readonly listItemAfter: number;
  /** Indent added per block-quote nesting level. */
  readonly quoteIndent: number;
  /** Space above/below a block quote. */
  readonly quoteSpacing: number;
  /** Space above/below a fenced code block (collapsed between its lines). */
  readonly codeSpacing: number;
  /** Space above/below a thematic break. */
  readonly ruleSpacing: number;
  /**
   * `w:spacing w:after` on the empty paragraph the renderer emits below every
   * table.
   *
   * `w:tbl` has no spacing property of its own, and two `w:tbl` elements with
   * nothing between them are one table as far as Word is concerned — so that
   * paragraph is load-bearing, not decorative. See `renderTable`.
   */
  readonly tableSpacing: number;
  /** Left/right padding inside a table cell. */
  readonly tableCellPaddingX: number;
  /** Top/bottom padding inside a table cell. */
  readonly tableCellPaddingY: number;
  /** Space above/below an image paragraph. */
  readonly figureSpacing: number;
  /** Border weight, eighth-points: the block-quote rule. */
  readonly quoteBorderSize: number;
  /** Gap between the block-quote rule and the text, in points. */
  readonly quoteBorderSpace: number;
  /** Border weight, eighth-points: the thematic break. */
  readonly ruleSize: number;
  /** Border weight, eighth-points: table grid lines. */
  readonly tableBorderSize: number;
}

/** Everything that distinguishes one heading level from the next. */
export interface HeadingSpec {
  /** Half-points. */
  readonly size: number;
  /** `RRGGBB`. */
  readonly color: string;
  readonly bold: boolean;
  readonly italics: boolean;
  /** Twips. */
  readonly spaceBefore: number;
  /** Twips. */
  readonly spaceAfter: number;
}

/** One rung of the bullet-glyph rotation. */
export interface BulletLevel {
  /** The glyph itself. Private-use code points require a matching `font`. */
  readonly text: string;
  /** Font the glyph is drawn in, or `null` to inherit the paragraph font. */
  readonly font: string | null;
}

/** Glyphs used for GFM task lists. */
export interface TaskGlyphs {
  /** `- [x]` */
  readonly checked: string;
  /** `- [ ]` */
  readonly unchecked: string;
  /** A non-checkbox item inside an otherwise-checkbox list. */
  readonly plain: string;
}

/* -------------------------------------------------------------------------- */
/* The theme                                                                   */
/* -------------------------------------------------------------------------- */

/** A fully resolved theme. Every field is present; nothing is optional. */
export interface Theme {
  readonly fonts: ThemeFonts;
  readonly colors: ThemeColors;
  readonly sizes: ThemeSizes;
  readonly spacing: ThemeSpacing;
  /** Per-level heading appearance, keyed by the model's `HeadingLevel`. */
  readonly headings: Readonly<Record<HeadingLevel, HeadingSpec>>;
  /**
   * Bullet glyphs, rotated by `level % bulletLevels.length`. The defaults are
   * the three Word itself writes (Symbol/Courier New/Wingdings), which every
   * major reader maps correctly.
   */
  readonly bulletLevels: readonly BulletLevel[];
  readonly taskGlyphs: TaskGlyphs;
  /** Number formats, rotated by `level % orderedFormats.length`. */
  readonly orderedFormats: readonly LevelFormatValue[];
  /** Colour used for the `==highlight==` mark. OOXML has a fixed palette. */
  readonly highlightColor: HighlightColorValue;
  /**
   * Maps a highlight scope to a colour. **The theme owns code colour.**
   *
   * Both of the renderer's highlighting paths read this map, which is what
   * makes `theme` mean something for a real code block:
   *
   *  - a {@link import("../model.js").CodeHighlightSpan} carried on the model;
   *    and
   *  - a {@link import("./types.js").HighlightSpan} from a
   *    {@link import("./types.js").Highlighter} that reports its scope (which
   *    `createHighlighter` does unless you hand it your own `scopeStyles`).
   *
   * Dotted scopes fall back to their prefix, so `title.function` resolves via
   * `title.function` -> `title` -> {@link ThemeColors.codeText}. A scope this
   * map does not name leaves the run at the `CodeBlock` style's own colour —
   * it does **not** fall through to a colour the span happened to carry, or a
   * palette could never take a scope back off an adapter.
   *
   * The corollary is that a palette must name every scope it wants coloured.
   * `tests/regressions.test.ts` runs the real `createHighlighter()` over a
   * fence in each of two dozen languages, collects every scope it actually
   * emits, and asserts that {@link DEFAULT_THEME}'s palette resolves all of
   * them — and that the ones `THEMES.print` leaves unresolved are exactly the
   * ones `PRINT_SCOPE_STYLES` documents as deliberately uncoloured. That is
   * what keeps "a fence in any supported language is themeable" true rather
   * than true-for-the-language-someone-happened-to-test.
   */
  readonly codePalette: Readonly<Record<string, string>>;
}

/** A partial theme. Merged one level deep over {@link DEFAULT_THEME}. */
export interface ThemeInit {
  readonly fonts?: Partial<ThemeFonts>;
  readonly colors?: Partial<ThemeColors>;
  readonly sizes?: Partial<ThemeSizes>;
  readonly spacing?: Partial<ThemeSpacing>;
  readonly headings?: Partial<Record<HeadingLevel, Partial<HeadingSpec>>>;
  readonly bulletLevels?: readonly BulletLevel[];
  readonly taskGlyphs?: Partial<TaskGlyphs>;
  readonly orderedFormats?: readonly LevelFormatValue[];
  readonly highlightColor?: HighlightColorValue;
  /**
   * **Replaces** {@link DEFAULT_THEME}'s palette rather than merging with it,
   * exactly like `bulletLevels` and `orderedFormats` above and like
   * `HighlighterOptions.scopeStyles` in `downword/highlight`.
   *
   * A palette is a designed set, not a bag of independent tokens: merging leaves
   * whichever scopes the new palette happens not to name painted in the old
   * one's inks, which is how `THEMES.print` used to ship eight screen colours
   * and miss the AAA floor its own docs promised. To adjust a few scopes,
   * spread: `{ codePalette: { ...DEFAULT_THEME.codePalette, keyword: "B00020" } }`.
   */
  readonly codePalette?: Readonly<Record<string, string>>;
}

/* -------------------------------------------------------------------------- */
/* Defaults                                                                    */
/* -------------------------------------------------------------------------- */

const HEADING_COLOR = "2F5496";
const HEADING_DEEP_COLOR = "1F3864";

/**
 * The default theme: Word's own Office style set for the built-in styles, and
 * a GitHub-flavoured palette for the things Word has no opinion about (code,
 * tables, quotes).
 */
export const DEFAULT_THEME: Theme = {
  fonts: {
    // Word's own pairing, current and previous: Microsoft 365 defaults to
    // Aptos / Aptos Display since 2024, and every build older than that (plus
    // LibreOffice, Pages and Google Docs) has Calibri / Calibri Light. Naming
    // both is what a `w:rFonts` can express of a fallback; see ThemeFontStack.
    body: { ascii: "Aptos", hAnsi: "Aptos", cs: "Calibri" },
    heading: { ascii: "Aptos Display", hAnsi: "Aptos Display", cs: "Calibri Light" },
    mono: "Consolas",
    monoFallback: "Courier New",
    symbol: "Segoe UI Symbol",
  },
  colors: {
    heading: HEADING_COLOR,
    headingDeep: HEADING_DEEP_COLOR,
    link: "0563C1",
    muted: "595959",
    quoteText: "404040",
    quoteBorder: "C8C8C8",
    codeText: "24292F",
    codeBackground: "F6F8FA",
    inlineCodeText: "24292F",
    inlineCodeBackground: "F0F1F3",
    rule: "BFBFBF",
    tableBorder: "BFBFBF",
    tableHeaderBackground: "F2F2F2",
  },
  sizes: {
    body: 22,
    title: 56,
    code: 20,
    footnote: 18,
    rule: 8,
  },
  spacing: {
    paragraphAfter: 160,
    line: 276,
    listIndent: 720,
    listHanging: 360,
    listItemAfter: 120,
    quoteIndent: 720,
    quoteSpacing: 120,
    codeSpacing: 120,
    ruleSpacing: 240,
    tableSpacing: 160,
    tableCellPaddingX: 108,
    tableCellPaddingY: 60,
    figureSpacing: 120,
    quoteBorderSize: 18,
    quoteBorderSpace: 12,
    ruleSize: 6,
    tableBorderSize: 4,
  },
  headings: {
    1: {
      size: 32,
      color: HEADING_COLOR,
      bold: true,
      italics: false,
      spaceBefore: 240,
      spaceAfter: 120,
    },
    2: {
      size: 28,
      color: HEADING_COLOR,
      bold: true,
      italics: false,
      spaceBefore: 200,
      spaceAfter: 100,
    },
    3: {
      size: 24,
      color: HEADING_DEEP_COLOR,
      bold: true,
      italics: false,
      spaceBefore: 180,
      spaceAfter: 80,
    },
    4: {
      size: 22,
      color: HEADING_COLOR,
      bold: true,
      italics: true,
      spaceBefore: 160,
      spaceAfter: 80,
    },
    5: {
      size: 22,
      color: HEADING_COLOR,
      bold: false,
      italics: false,
      spaceBefore: 160,
      spaceAfter: 80,
    },
    6: {
      size: 22,
      color: HEADING_DEEP_COLOR,
      bold: false,
      italics: true,
      spaceBefore: 160,
      spaceAfter: 80,
    },
  },
  bulletLevels: [
    // U+F0B7 in Symbol renders as a solid bullet; this is byte-for-byte what
    // Word writes for a default bulleted list.
    { text: "", font: "Symbol" },
    { text: "o", font: "Courier New" },
    // U+F0A7 in Wingdings renders as a small filled square.
    { text: "", font: "Wingdings" },
  ],
  taskGlyphs: {
    // BALLOT BOX WITH CHECK / BALLOT BOX. Real Unicode rather than a
    // private-use Wingdings code point, so non-Word readers show a checkbox.
    checked: "☑",
    unchecked: "☐",
    plain: "•",
  },
  orderedFormats: ["decimal", "lowerLetter", "lowerRoman"],
  highlightColor: "yellow",
  // GitHub Light, measured against `colors.codeBackground` (F6F8FA) rather than
  // against white, because that shading is what these inks are actually painted
  // on. The floor is 4.74:1 (`8250DF`), so every scope clears WCAG **AA** for
  // normal-size text. It does not clear AAA and is not meant to - that is what
  // `THEMES.print` is for; see `highlight/palette.ts`.
  codePalette: {
    keyword: "CF222E",
    built_in: "8250DF",
    type: "953800",
    literal: "0550AE",
    number: "0550AE",
    operator: "0550AE",
    punctuation: "24292F",
    property: "0550AE",
    regexp: "0A3069",
    string: "0A3069",
    char: "0A3069",
    subst: "24292F",
    symbol: "953800",
    variable: "953800",
    "variable.language": "CF222E",
    "template-variable": "CF222E",
    title: "8250DF",
    "title.function": "8250DF",
    "title.class": "953800",
    section: "0550AE",
    params: "24292F",
    // The legacy wrapper scope a few grammars (php, go) put round a whole
    // signature. Everything that matters inside it carries `title.function`,
    // `params` or `type`, and the emitter keeps only the innermost scope - so
    // what actually reaches a run scoped plain `function` is the punctuation
    // and whitespace between them. Hence the punctuation ink rather than the
    // callable one; measured, not assumed (`tests/regressions.test.ts`).
    function: "24292F",
    // 656D76, not GitHub's screen 6E7781: that is 4.27:1 on the F6F8FA
    // shading this text sits on, i.e. below AA. This is 4.93:1.
    comment: "656D76",
    doctag: "CF222E",
    meta: "0550AE",
    tag: "116329",
    name: "116329",
    attr: "0550AE",
    attribute: "0550AE",
    // highlight.js's CSS/SCSS/LESS scopes are dash-separated, not dotted, so
    // `scopeColor`'s prefix walk cannot reach `selector` from them. Without
    // these five entries an entire stylesheet is unthemeable - which is exactly
    // what happened while a highlighter's own inks were still a fallback.
    selector: "116329",
    "selector-tag": "116329",
    "selector-pseudo": "116329",
    "selector-class": "0550AE",
    "selector-id": "0550AE",
    "selector-attr": "0550AE",
    bullet: "0550AE",
    quote: "656D76",
    // Markdown-inside-code: a fenced block or an inline span in a ```md fence.
    code: "656D76",
    formula: "656D76",
    link: "0A3069",
    emphasis: "24292F",
    strong: "24292F",
    addition: "116329",
    deletion: "82071E",
  },
};

/* -------------------------------------------------------------------------- */
/* Resolution                                                                  */
/* -------------------------------------------------------------------------- */

function mergeHeading(base: HeadingSpec, init: Partial<HeadingSpec> | undefined): HeadingSpec {
  return init === undefined ? base : { ...base, ...init };
}

/**
 * Merges a partial theme over {@link DEFAULT_THEME}.
 *
 * Idempotent: `resolveTheme(resolveTheme(x))` deep-equals `resolveTheme(x)`,
 * which is what lets {@link Theme} be passed straight back into
 * `buildStyles()`.
 */
export function resolveTheme(init: ThemeInit = {}): Theme {
  const headings = init.headings;
  return {
    fonts: { ...DEFAULT_THEME.fonts, ...init.fonts },
    colors: { ...DEFAULT_THEME.colors, ...init.colors },
    sizes: { ...DEFAULT_THEME.sizes, ...init.sizes },
    spacing: { ...DEFAULT_THEME.spacing, ...init.spacing },
    headings: {
      1: mergeHeading(DEFAULT_THEME.headings[1], headings?.[1]),
      2: mergeHeading(DEFAULT_THEME.headings[2], headings?.[2]),
      3: mergeHeading(DEFAULT_THEME.headings[3], headings?.[3]),
      4: mergeHeading(DEFAULT_THEME.headings[4], headings?.[4]),
      5: mergeHeading(DEFAULT_THEME.headings[5], headings?.[5]),
      6: mergeHeading(DEFAULT_THEME.headings[6], headings?.[6]),
    },
    bulletLevels:
      init.bulletLevels !== undefined && init.bulletLevels.length > 0
        ? init.bulletLevels
        : DEFAULT_THEME.bulletLevels,
    taskGlyphs: { ...DEFAULT_THEME.taskGlyphs, ...init.taskGlyphs },
    orderedFormats:
      init.orderedFormats !== undefined && init.orderedFormats.length > 0
        ? init.orderedFormats
        : DEFAULT_THEME.orderedFormats,
    highlightColor: init.highlightColor ?? DEFAULT_THEME.highlightColor,
    codePalette: init.codePalette ?? DEFAULT_THEME.codePalette,
  };
}

/* -------------------------------------------------------------------------- */
/* Presets                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * ## What a preset may and may not change
 *
 * A theme is meant to be a `styles.xml` swap and nothing else: the same
 * markdown under two presets must produce the *same document* differently
 * dressed, so that everything Word does with a style — Design → Style Set, the
 * Navigation pane, "update Heading 2 everywhere" — keeps working, and so that
 * two conversions can be diffed against each other.
 *
 * Most theme tokens honour that for free, because `styles.ts` is the only
 * reader. A short list does not, because OOXML has no style that could hold
 * them, and `render/blocks.ts` / `render/inline.ts` therefore stamp them onto
 * elements in `word/document.xml` (or into `word/numbering.xml`):
 *
 * | token | where it lands | why it cannot be a style |
 * | --- | --- | --- |
 * | `spacing.listIndent`, `spacing.listHanging` | `w:ind` on every list paragraph, and `numbering.xml` | the indent depends on the item's level |
 * | `spacing.listItemAfter` | `w:spacing` on a *loose* list's items | looseness is a property of the list, not the style |
 * | `spacing.quoteIndent` | `w:ind` on a nested quote | depends on nesting depth |
 * | `colors.quoteBorder`, `spacing.quoteBorderSize`, `spacing.quoteBorderSpace` | `w:pBdr` on a nested quote | ditto |
 * | `sizes.body` | table column widths (`w:gridCol`, `w:tcW`) | a column width is a measurement, not a style |
 * | `spacing.tableCellPadding*`, `colors.tableBorder`, `spacing.tableBorderSize`, `colors.tableHeaderBackground` | `w:tblPr` / `w:tcPr` | downword emits no table *style* |
 * | `taskGlyphs.*`, `fonts.symbol`, `bulletLevels`, `orderedFormats` | `numbering.xml`, and the task-list glyph run | list markers live in the numbering part |
 * | `highlightColor` | `w:highlight` on a `==marked==` run | OOXML's fixed 17-colour palette, no style |
 * | `codePalette` | `w:color` per highlighted span | a span's scope is content, not style |
 * | `colors.link` | `w:color` on a run that is *both* inline code and a link | one `w:rStyle` per run, and `CodeChar` took the slot |
 *
 * So the three presets below hold every one of those constant, with two
 * deliberate exceptions that are documented, measured and asserted in
 * `tests/themes.test.ts`:
 *
 *  - **`academic` changes `sizes.body`** (22 → 24 half-points). Twelve-point
 *    Times is the entire point of the preset, and the column heuristic in
 *    `computeColumnWidths` reads the body size, so a document *with a table*
 *    differs from `default` in its `w:gridCol`/`w:tcW` values and in nothing
 *    else.
 *  - **every preset changes `colors.link`.** It reaches `document.xml` only for
 *    the one construct that is both inline code and a link; ordinary links take
 *    it from the `Hyperlink` character style, which is where a link colour
 *    belongs. Fixing that would mean teaching `render/inline.ts` to compose the
 *    two, not freezing a colour every theme wants to own.
 *
 * The code inks are the third thing held constant, for a different reason:
 * `codePalette` and `colors.code*` are a *measured pair* (every ink clears WCAG
 * AA against `codeBackground`), so a preset that moved the shading would
 * silently invalidate a contrast guarantee the docs make. Swapping the palette
 * is a separate axis — `{ ...THEMES.academic, codePalette: THEMES.print.codePalette }`
 * composes the two.
 */

/**
 * GitHub-flavoured: the system sans, near-black headings, tighter leading.
 *
 * What a README looks like on github.com, as far as a Word document can: the
 * heading colour is the body colour (GitHub does not tint headings), the
 * leading is single rather than Word's 1.15, and paragraph spacing is smaller
 * throughout. Segoe UI is the Windows system face at the head of GitHub's own
 * font stack; Arial is the fallback, being the one sans every reader has.
 */
export const GITHUB_THEME: Theme = resolveTheme({
  fonts: {
    body: fontStack("Segoe UI", "Arial"),
    heading: fontStack("Segoe UI", "Arial"),
  },
  colors: {
    // GitHub Primer: fg.default for headings, fg.muted for the quiet text.
    heading: "1F2328",
    headingDeep: "1F2328",
    link: "0969DA",
    muted: "59636E",
    quoteText: "59636E",
    rule: "D1D9E0",
  },
  sizes: {
    // `body` is deliberately Word's 11pt: see the note above on why the body
    // size is the one size a preset cannot move without moving table columns.
    title: 48,
  },
  spacing: {
    paragraphAfter: 120,
    line: 240,
    quoteSpacing: 100,
    codeSpacing: 100,
    ruleSpacing: 200,
    tableSpacing: 120,
    figureSpacing: 100,
  },
  headings: {
    1: { size: 32, color: "1F2328", spaceBefore: 320, spaceAfter: 80 },
    2: { size: 26, color: "1F2328", spaceBefore: 280, spaceAfter: 80 },
    3: { size: 24, color: "1F2328", spaceBefore: 240, spaceAfter: 60 },
    4: { size: 22, color: "1F2328", bold: true, italics: false, spaceBefore: 240, spaceAfter: 60 },
    5: { size: 22, color: "1F2328", bold: true, spaceBefore: 240, spaceAfter: 60 },
    6: { size: 22, color: "59636E", bold: true, italics: false, spaceBefore: 240, spaceAfter: 60 },
  },
});

/**
 * Times New Roman at 12 pt, 1.5 leading, black headings.
 *
 * The manuscript default: no colour anywhere a printer would have to render in
 * grey, a serif body, and Courier New for code (the face a monospace listing
 * has had in a typeset paper for forty years) with Consolas as the fallback.
 * Cambria is the body fallback because it is the serif every Office install
 * has, including the ones with no Times New Roman licence.
 *
 * See {@link ACADEMIC_DOUBLE_THEME} for the double-spaced variant.
 */
export const ACADEMIC_THEME: Theme = resolveTheme({
  fonts: {
    body: fontStack("Times New Roman", "Cambria"),
    heading: fontStack("Times New Roman", "Cambria"),
    mono: "Courier New",
    monoFallback: "Consolas",
  },
  colors: {
    heading: "000000",
    headingDeep: "000000",
    // Dark enough to read as ink on paper, blue enough to read as a link.
    link: "1F4E79",
    muted: "3C3C3C",
    quoteText: "262626",
    rule: "808080",
  },
  sizes: {
    // 12 pt. The one document.xml-reaching token a preset moves; see above.
    body: 24,
    title: 36,
    code: 22,
    footnote: 20,
  },
  spacing: {
    paragraphAfter: 120,
    // 1.5 lines. `ACADEMIC_DOUBLE_THEME` is this and 480.
    line: 360,
  },
  headings: {
    1: { size: 28, color: "000000", spaceBefore: 240, spaceAfter: 120 },
    2: { size: 26, color: "000000", spaceBefore: 240, spaceAfter: 120 },
    3: { size: 24, color: "000000", spaceBefore: 240, spaceAfter: 120 },
    4: { size: 24, color: "000000", bold: false, italics: true, spaceBefore: 240, spaceAfter: 120 },
    5: { size: 24, color: "000000", bold: false, italics: true, spaceBefore: 240, spaceAfter: 120 },
    6: { size: 24, color: "000000", bold: false, italics: true, spaceBefore: 240, spaceAfter: 120 },
  },
});

/**
 * {@link ACADEMIC_THEME}, double-spaced.
 *
 * The submission format most journals and every thesis office ask for, and the
 * reason it is a preset rather than an option: `line: 480` with
 * `lineRule="auto"` is *the* difference — one token, in `styles.xml`, reaching
 * `docDefaults` and every style that restates the leading. Nothing else moves,
 * which is exactly what a reader who switches between the two should see.
 */
export const ACADEMIC_DOUBLE_THEME: Theme = {
  ...ACADEMIC_THEME,
  spacing: { ...ACADEMIC_THEME.spacing, line: 480 },
};

/**
 * Resolves a highlight scope to a colour, walking `a.b.c` -> `a.b` -> `a`.
 *
 * Returns `null` when nothing matches, which the renderer reads as "leave the
 * run at the code block's default colour".
 */
export function scopeColor(theme: Theme, scope: string): string | null {
  let key = scope;
  for (;;) {
    const found = theme.codePalette[key];
    if (found !== undefined) return found;
    const dot = key.lastIndexOf(".");
    if (dot <= 0) return null;
    key = key.slice(0, dot);
  }
}
