/**
 * One light syntax theme, tuned for **paper**.
 *
 * A .docx is a print artefact before it is a screen artefact: it gets emailed,
 * PDF'd and stapled. Every terminal theme is therefore disqualified — a dark
 * theme in Word means either unreadable dark-on-white text or a full-bleed dark
 * rectangle that eats a printer cartridge. So: dark ink on the light `CodeBlock`
 * shading, and enough contrast that a grayscale laser printer still resolves it.
 *
 * ## Two constraints the screen never imposes
 *
 * **1. Contrast against the shading, not against white.** The renderer paints
 * `CodeBlock` with `#F6F8FA`, which is *lighter* than mid-grey but still eats
 * ~6% of the available luminance range. Every ink below is measured against
 * that shading (the real background) as well as against bare white (what a
 * reader gets if they restyle the document or print without background
 * shading). The numbers are WCAG 2.x contrast ratios:
 *
 * | role       | ink      | on `F6F8FA` | on `FFFFFF` |
 * | ---------- | -------- | ----------- | ----------- |
 * | `text`     | `24292F` | 13.76       | 14.65       |
 * | `keyword`  | `A40E26` |  7.39       |  7.87       |
 * | `string`   | `0A3069` | 12.03       | 12.81       |
 * | `comment`  | `4B535D` |  7.32       |  7.79       |
 * | `number`   | `0550AE` |  7.13       |  7.59       |
 * | `type`     | `8A3400` |  7.67       |  8.17       |
 * | `callable` | `5A32A3` |  8.15       |  8.68       |
 * | `markup`   | `0E5427` |  8.52       |  9.07       |
 * | `meta`     | `044289` |  9.19       |  9.79       |
 * | `removed`  | `82071E` |  9.87       | 10.51       |
 *
 * The floor is 7.13:1, so every ink clears WCAG **AAA** (7:1) for normal-size
 * text — not merely AA. That headroom is deliberate: code is set at 10 pt, it
 * is often photocopied, and the shading itself is not guaranteed to survive.
 *
 * **2. Hue is not available in grayscale.** Printed monochrome, `keyword`
 * (7.39) and `comment` (7.32) are the same grey. Two scopes therefore carry a
 * non-colour signal as well: keywords are **bold** and comments are *italic*.
 * That is the classic typeset-listing convention (LaTeX `listings` does the
 * same), and it means a black-and-white printout still separates the five
 * things a reader scans for: prose-grey italic comments, bold keywords, near-
 * black identifiers, dark strings, mid-tone literals.
 *
 * Nothing here is baked in — pass {@link HighlighterOptions.scopeStyles} to
 * replace the map wholesale, or spread {@link PRINT_SCOPE_STYLES} and override
 * a few scopes.
 */

/** Direct run formatting for one highlight scope. */
export interface ScopeStyle {
  /** `RRGGBB`, no `#`. Omitted means "inherit the CodeBlock style's colour". */
  readonly color?: string | undefined;
  readonly bold?: boolean | undefined;
  readonly italic?: boolean | undefined;
}

/* -------------------------------------------------------------------------- */
/* Inks                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * The ten inks, by role. `RRGGBB`, no `#`.
 *
 * Ten is a deliberate ceiling: a printed page with more than about eight
 * distinguishable text colours reads as noise, and OOXML gives no way to tune
 * them per reader. Scopes are mapped *onto* these roles in
 * {@link PRINT_SCOPE_STYLES} rather than each getting its own colour.
 */
export const PRINT_INK = {
  /**
   * Identifiers, operators, punctuation — everything unremarkable.
   *
   * Matches the renderer's `colors.codeText` default, which is why
   * {@link PRINT_SCOPE_STYLES} leaves ordinary text *uncoloured* rather than
   * stamping this value onto every run: the `CodeBlock` style already says it.
   */
  text: "24292F",
  /** Language keywords, `type`, doctags, template tags. Set bold as well. */
  keyword: "A40E26",
  /** String, regexp and character literals; also link targets in markup. */
  string: "0A3069",
  /** Comments, quotes, formulae. Set italic as well. */
  comment: "4B535D",
  /** Numbers, booleans, attributes, operators-as-values, plain variables. */
  number: "0550AE",
  /** Built-ins, symbols, list bullets — the "provided by the platform" tone. */
  type: "8A3400",
  /** Function and class names at their definition or call site. */
  callable: "5A32A3",
  /** Tag/selector names and diff additions: structure rather than value. */
  markup: "0E5427",
  /** Preprocessor lines, shebangs, decorators, annotations. */
  meta: "044289",
  /** Diff deletions. Distinct from `keyword` red by being much darker. */
  removed: "82071E",
} as const satisfies Readonly<Record<string, string>>;

/** The role names of {@link PRINT_INK}. */
export type PrintInkRole = keyof typeof PRINT_INK;

/* -------------------------------------------------------------------------- */
/* Scope map                                                                   */
/* -------------------------------------------------------------------------- */

const KEYWORD: ScopeStyle = { color: PRINT_INK.keyword, bold: true };
const STRING: ScopeStyle = { color: PRINT_INK.string };
const COMMENT: ScopeStyle = { color: PRINT_INK.comment, italic: true };
const NUMBER: ScopeStyle = { color: PRINT_INK.number };
const TYPE: ScopeStyle = { color: PRINT_INK.type };
const CALLABLE: ScopeStyle = { color: PRINT_INK.callable };
const MARKUP: ScopeStyle = { color: PRINT_INK.markup };
const META: ScopeStyle = { color: PRINT_INK.meta };

/**
 * highlight.js scope -> run formatting.
 *
 * The scope names are highlight.js's own (`hljs-` prefix already stripped by
 * the emitter). Dotted scopes resolve longest-prefix-first via
 * {@link resolveScopeStyle}, so `title.function.invoke` finds `title.function`
 * and `meta.string` finds `meta.string` before it would find `meta`.
 *
 * The grouping follows GitHub Light's structure — that mapping is well tested
 * against ~200 grammars — with this file's print inks substituted for its
 * screen ones. Scopes deliberately absent (`punctuation`, `operator`, `subst`,
 * `params`, `tag`, `template-tag`, and the legacy `function` wrapper — which,
 * once the emitter has kept only the innermost scope, carries nothing but the
 * punctuation between a signature's parts) fall through to the `CodeBlock`
 * style's own colour, which keeps the XML small and the page calm.
 */
export const PRINT_SCOPE_STYLES: Readonly<Record<string, ScopeStyle>> = {
  /* Keywords and the things that behave like them. */
  keyword: KEYWORD,
  doctag: KEYWORD,
  type: KEYWORD,
  "template-variable": KEYWORD,
  "variable.language": KEYWORD,
  "variable.constant": NUMBER,

  /* Names being defined or invoked. */
  title: CALLABLE,
  "title.class": CALLABLE,
  "title.class.inherited": CALLABLE,
  "title.function": CALLABLE,
  "title.function.invoke": CALLABLE,

  /* Values. */
  literal: NUMBER,
  number: NUMBER,
  variable: NUMBER,
  property: NUMBER,
  attr: NUMBER,
  attribute: NUMBER,
  "selector-attr": NUMBER,
  "selector-class": NUMBER,
  "selector-id": NUMBER,

  /* Text-shaped literals. */
  string: STRING,
  regexp: STRING,
  char: STRING,
  "char.escape": STRING,
  link: STRING,
  "meta.string": STRING,

  /* Platform-provided names. */
  built_in: TYPE,
  symbol: TYPE,
  bullet: TYPE,

  /* Prose inside code. */
  comment: COMMENT,
  quote: COMMENT,
  code: COMMENT,
  formula: COMMENT,

  /* Markup structure. */
  name: MARKUP,
  "selector-tag": MARKUP,
  "selector-pseudo": MARKUP,
  addition: MARKUP,
  deletion: { color: PRINT_INK.removed },
  section: { color: PRINT_INK.number, bold: true },

  /* Out-of-band lines: `#include`, `#!/bin/sh`, `@Override`, `>>> `. */
  meta: META,
  "meta.prompt": META,
  "meta.keyword": META,

  /* Markdown emphasis carried through a highlighted block. */
  emphasis: { italic: true },
  strong: { bold: true },
};

/**
 * The colour half of {@link PRINT_SCOPE_STYLES}, shaped for the renderer's
 * `theme.codePalette`.
 *
 * The renderer has a *second* highlighting path: a `CodeBlockNode` may arrive
 * carrying `highlights` (offset + scope-name spans) instead of going through a
 * {@link import("../render/types.js").Highlighter}, and those spans are
 * coloured from `theme.codePalette`. Feeding this constant to `resolveTheme()`
 * makes both paths render identically:
 *
 * ```ts
 * renderDocument(doc, { theme: { codePalette: PRINT_CODE_PALETTE } });
 * ```
 *
 * Bold/italic are dropped, because a theme palette maps scopes to colours only.
 */
export const PRINT_CODE_PALETTE: Readonly<Record<string, string>> = Object.freeze(
  Object.fromEntries(
    Object.entries(PRINT_SCOPE_STYLES).flatMap(([scope, style]) =>
      style.color === undefined ? [] : [[scope, style.color] as const],
    ),
  ),
);

/* -------------------------------------------------------------------------- */
/* Resolution                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Resolves a scope to a style, walking `a.b.c` -> `a.b` -> `a`.
 *
 * highlight.js emits *tiered* scopes, and a grammar is free to invent
 * `title.function.something-new`; falling back along the dots means an unknown
 * leaf still picks up its family's colour instead of rendering plain. Mirrors
 * `scopeColor()` in the renderer's theme so the two paths agree.
 *
 * Returns `null` when nothing in the family matches, which callers read as
 * "leave the run at the code block's default colour".
 */
export function resolveScopeStyle(
  styles: Readonly<Record<string, ScopeStyle>>,
  scope: string,
): ScopeStyle | null {
  let key = scope;
  for (;;) {
    const found = styles[key];
    if (found !== undefined) return found;
    const dot = key.lastIndexOf(".");
    if (dot <= 0) return null;
    key = key.slice(0, dot);
  }
}
