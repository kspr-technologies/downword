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

/** Typefaces used across the document. */
export interface ThemeFonts {
  /** Body text; becomes `docDefaults`' `w:rFonts`. */
  readonly body: string;
  /** Headings and the title. */
  readonly heading: string;
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
    body: "Calibri",
    heading: "Calibri Light",
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
