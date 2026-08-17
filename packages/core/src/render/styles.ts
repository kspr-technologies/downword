/**
 * `word/styles.xml` generation.
 *
 * ## Why this file exists at all
 *
 * A renderer can make a heading look like a heading in two ways: stamp
 * `<w:b/><w:sz w:val="32"/><w:color w:val="2F5496"/>` onto the run (direct
 * formatting), or point the paragraph at a *style* with `<w:pStyle
 * w:val="Heading1"/>` and define that style once. Only the second survives
 * contact with a real user:
 *
 *  - Word's **Design -> Style Set** and theme pickers rewrite style
 *    *definitions*. Direct formatting is immune to them, so a document built
 *    from direct formatting cannot be restyled at all.
 *  - The Navigation pane, `TOC` fields, "select all text with similar
 *    formatting", accessibility checkers and PDF tagging all key off
 *    `Heading1..6` / `Title` — not off font size.
 *
 * So every visual decision that OOXML can express as a style property is
 * emitted here, and the body only ever references style ids. The handful of
 * properties that are genuinely positional — how deep *this* quote is nested,
 * how wide *this* table column is — stay on the element, because there is no
 * style that could hold them.
 *
 * ## Which ids are built in
 *
 * `Title`, `Heading1..6`, `Strong`, `ListParagraph`, `Hyperlink`,
 * `FootnoteReference`, `FootnoteText` and `FootnoteTextChar` are all keys of
 * docx's `styles.default`, i.e. Word built-ins whose style *ids* Word already
 * recognises. `Quote`, `Normal` and `DefaultParagraphFont` are **not**
 * `styles.default` keys even though Word treats all three as built-ins, so they
 * are declared by hand with an explicit `id`/`name`. Everything else
 * (`CodeBlock`, `CodeChar`, …) is ours, named so that it reads sensibly in
 * Word's style gallery.
 *
 * `Normal` and `DefaultParagraphFont` are the roots: every `basedOn` in this
 * file resolves to a style that is actually in the part, and every plain body
 * paragraph carries `<w:pStyle w:val="Normal"/>`. Without them the twenty-three
 * `basedOn` references dangled and body text was the one thing in the document
 * a Style Set could not reach — the exact claim the section above makes.
 *
 * Built-in `w:name` values are spelled exactly as Word writes them —
 * `"heading 1"`, lowercase, not `"Heading 1"` — because that is how a reader
 * matches a style to its built-in identity when the style set changes.
 */

import {
  AlignmentType,
  BorderStyle,
  LineRuleType,
  ShadingType,
  UnderlineType,
  type IBaseParagraphStyleOptions,
  type ICharacterStyleOptions,
  type IFontAttributesProperties,
  type IParagraphStyleOptions,
  type IStylesOptions,
} from "docx";

import type { HeadingLevel } from "../model.js";
import { resolveTheme, type Theme, type ThemeInit } from "./theme.js";

/* -------------------------------------------------------------------------- */
/* Style ids                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Every style id the renderer emits.
 *
 * Exported so that callers (and the golden tests) can assert on ids instead of
 * re-typing string literals.
 */
export const STYLE_IDS = {
  /**
   * Built-in. The root of the style tree and the style of ordinary body text.
   *
   * Word invents a `Normal` when a document does not declare one, so nothing
   * *broke* while this was missing — but a `w:basedOn` pointing at a style that
   * is not in the part is ignored, and a Style Set can only rewrite styles that
   * exist. Every other style here is `basedOn` this one, and every plain
   * paragraph references it, so Design -> Style Set reaches the body text too.
   */
  normal: "Normal",
  /** Built-in. The root character style; every character style is based on it. */
  defaultParagraphFont: "DefaultParagraphFont",
  /** Built-in. Used for the frontmatter title block. */
  title: "Title",
  /** Built-in. Indexed by heading level, 1-based. */
  heading: ["Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6"],
  /** Built-in id, but *not* a docx `styles.default` key — declared by hand. */
  quote: "Quote",
  /** Built-in. Every list paragraph carries it. */
  listParagraph: "ListParagraph",
  /** Built-in character style for hyperlink text. */
  hyperlink: "Hyperlink",
  /** Built-in character style for the superscript footnote marker. */
  footnoteReference: "FootnoteReference",
  /** Built-in paragraph style for footnote bodies. */
  footnoteText: "FootnoteText",
  /** Ours: inline `` `code` ``. */
  codeChar: "CodeChar",
  /** Ours: one paragraph per line of a fenced code block. */
  codeBlock: "CodeBlock",
  /** Ours: raw block-level HTML under the `"raw"` policy. */
  htmlBlock: "HtmlBlock",
  /** Ours: raw inline HTML under the `"raw"` policy. */
  htmlChar: "HtmlChar",
  /** Ours: the empty bottom-bordered paragraph a `---` becomes. */
  horizontalRule: "HorizontalRule",
  /** Ours: body cells. */
  tableText: "TableText",
  /** Ours: header cells. */
  tableHeading: "TableHeading",
  /** Ours: the empty paragraph that closes a table. See `renderTable`. */
  tableSpacing: "TableSpacing",
  /** Ours: a paragraph whose only content is an image. */
  figure: "Figure",
  /** Ours: a paragraph standing in for an image that could not be resolved. */
  imagePlaceholder: "ImagePlaceholder",
  /** Ours: an inline image that could not be resolved. */
  imagePlaceholderChar: "ImagePlaceholderChar",
} as const;

/** Returns the built-in style id for a heading level. */
export function headingStyleId(level: HeadingLevel): string {
  return STYLE_IDS.heading[level - 1] ?? "Heading6";
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * The monospace fallback chain.
 *
 * OOXML has no font stack, but it does have four *scripts* per run. Putting the
 * preferred face on `ascii`/`hAnsi` and the ubiquitous one on `cs` gives
 * readers that cannot resolve Consolas something metric-appropriate to fall
 * back to, instead of dropping to a proportional default. `eastAsia` is left
 * unset on purpose so CJK inside a code block keeps using a CJK-capable face.
 */
function monoFont(theme: Theme): IFontAttributesProperties {
  return {
    ascii: theme.fonts.mono,
    hAnsi: theme.fonts.mono,
    cs: theme.fonts.monoFallback,
  };
}

/**
 * A font size, stated for both the Latin and the complex-script run.
 *
 * OOXML sizes a run twice: `w:sz` governs the `w:ascii`/`w:hAnsi` text and
 * `w:szCs` governs anything the layout engine classifies as complex script —
 * Arabic, Hebrew, Thai, Devanagari. With `w:szCs` unset, Word falls back to its
 * own default (10 pt) for that text, so an Arabic heading renders at body size
 * next to its Latin siblings. Word always writes both; so do we.
 */
function sized(halfPoints: number): { size: number; sizeComplexScript: number } {
  return { size: halfPoints, sizeComplexScript: halfPoints };
}

function headingStyle(theme: Theme, level: HeadingLevel): IBaseParagraphStyleOptions {
  const spec = theme.headings[level];
  return {
    name: `heading ${level}`,
    basedOn: STYLE_IDS.normal,
    next: STYLE_IDS.normal,
    quickFormat: true,
    uiPriority: 9,
    paragraph: {
      spacing: {
        before: spec.spaceBefore,
        after: spec.spaceAfter,
        line: theme.spacing.line,
        lineRule: LineRuleType.AUTO,
      },
      // Headings must not be orphaned at the foot of a page, and must show up
      // in the Navigation pane / TOC at the right depth.
      keepNext: true,
      keepLines: true,
      outlineLevel: level - 1,
    },
    run: {
      font: theme.fonts.heading,
      ...sized(spec.size),
      bold: spec.bold,
      italics: spec.italics,
      color: spec.color,
    },
  };
}

/* -------------------------------------------------------------------------- */
/* buildStyles                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Builds the complete `styles.xml` payload for a theme.
 *
 * Accepts a partial theme so it is usable standalone (`buildStyles({ colors:
 * { link: "AA0000" } })`); `renderDocument()` passes an already-resolved
 * {@link Theme}, which round-trips unchanged because {@link resolveTheme} is
 * idempotent.
 */
export function buildStyles(theme: ThemeInit | Theme = {}): IStylesOptions {
  const t = resolveTheme(theme);
  const mono = monoFont(t);

  const paragraphStyles: readonly IParagraphStyleOptions[] = [
    {
      // The root of the tree. It deliberately declares no properties: every
      // visual default already lives in `docDefaults` below, and duplicating
      // them here would give Word two places to disagree. What it *does* buy is
      // the sixteen `w:basedOn w:val="Normal"` references in this file
      // resolving to something, and a `w:pStyle` for body text that Design ->
      // Style Set can rewrite.
      //
      // (`w:default="1"` is what would make Word apply it to a paragraph with
      // no `w:pStyle` at all; docx's `IStyleOptions` cannot express the
      // attribute, so the renderer references the id explicitly instead.)
      id: STYLE_IDS.normal,
      name: "Normal",
      quickFormat: true,
      uiPriority: 0,
    },
    {
      // Word knows this id. Declared here rather than via `styles.default`
      // because docx's IDefaultStylesOptions has no `quote` key.
      id: STYLE_IDS.quote,
      name: "Quote",
      basedOn: STYLE_IDS.normal,
      next: STYLE_IDS.normal,
      quickFormat: true,
      uiPriority: 29,
      paragraph: {
        spacing: {
          before: t.spacing.quoteSpacing,
          after: t.spacing.quoteSpacing,
          line: t.spacing.line,
          lineRule: LineRuleType.AUTO,
        },
        border: {
          left: {
            style: BorderStyle.SINGLE,
            color: t.colors.quoteBorder,
            size: t.spacing.quoteBorderSize,
            space: t.spacing.quoteBorderSpace,
          },
        },
      },
      run: { color: t.colors.quoteText },
    },
    {
      id: STYLE_IDS.codeBlock,
      name: "Code Block",
      basedOn: STYLE_IDS.normal,
      // A code block is a run of consecutive CodeBlock paragraphs, so the
      // style points at itself: pressing Enter inside one stays in the block.
      next: STYLE_IDS.codeBlock,
      uiPriority: 30,
      paragraph: {
        shading: { type: ShadingType.CLEAR, color: "auto", fill: t.colors.codeBackground },
        spacing: {
          before: t.spacing.codeSpacing,
          after: t.spacing.codeSpacing,
          line: 240,
          lineRule: LineRuleType.AUTO,
        },
        // The whole reason `before`/`after` can be non-zero: contextualSpacing
        // suppresses them *between* paragraphs of the same style, so the block
        // is airtight inside and breathes at its edges. No per-line direct
        // spacing is needed, and no first/last special case exists.
        contextualSpacing: true,
        keepLines: true,
      },
      run: { font: mono, ...sized(t.sizes.code), color: t.colors.codeText },
    },
    {
      id: STYLE_IDS.htmlBlock,
      name: "Raw HTML Block",
      basedOn: STYLE_IDS.codeBlock,
      next: STYLE_IDS.normal,
      uiPriority: 31,
      run: { font: mono, ...sized(t.sizes.code), color: t.colors.muted },
    },
    {
      id: STYLE_IDS.horizontalRule,
      name: "Horizontal Rule",
      basedOn: STYLE_IDS.normal,
      next: STYLE_IDS.normal,
      uiPriority: 32,
      paragraph: {
        border: {
          bottom: {
            style: BorderStyle.SINGLE,
            color: t.colors.rule,
            size: t.spacing.ruleSize,
            space: 1,
          },
        },
        spacing: {
          before: t.spacing.ruleSpacing,
          after: t.spacing.ruleSpacing,
          line: 240,
          lineRule: LineRuleType.AUTO,
        },
      },
      // The paragraph is empty, so its mark alone sets the line height; a small
      // size keeps the rule from reserving a full text line.
      run: sized(t.sizes.rule),
    },
    {
      id: STYLE_IDS.tableText,
      name: "Table Text",
      basedOn: STYLE_IDS.normal,
      next: STYLE_IDS.tableText,
      uiPriority: 33,
      paragraph: {
        spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      },
    },
    {
      id: STYLE_IDS.tableHeading,
      name: "Table Heading",
      basedOn: STYLE_IDS.tableText,
      next: STYLE_IDS.tableText,
      uiPriority: 34,
      run: { bold: true },
    },
    {
      // The empty paragraph below every table. Its own height is one small
      // paragraph mark, so what the reader sees is `after` - and what Word sees
      // is the block-level element that keeps two consecutive `w:tbl` elements
      // from being merged into one table. See `renderTable`.
      id: STYLE_IDS.tableSpacing,
      name: "Table Spacing",
      basedOn: STYLE_IDS.normal,
      next: STYLE_IDS.normal,
      // Word hides styles nobody should apply by hand from the gallery.
      semiHidden: true,
      unhideWhenUsed: true,
      uiPriority: 99,
      paragraph: {
        spacing: {
          before: 0,
          after: t.spacing.tableSpacing,
          line: 240,
          lineRule: LineRuleType.AUTO,
        },
      },
      run: sized(t.sizes.rule),
    },
    {
      id: STYLE_IDS.figure,
      name: "Figure",
      basedOn: STYLE_IDS.normal,
      next: STYLE_IDS.normal,
      uiPriority: 35,
      paragraph: {
        alignment: AlignmentType.CENTER,
        spacing: { before: t.spacing.figureSpacing, after: t.spacing.figureSpacing },
      },
    },
    {
      id: STYLE_IDS.imagePlaceholder,
      name: "Image Placeholder",
      basedOn: STYLE_IDS.figure,
      next: STYLE_IDS.normal,
      uiPriority: 36,
      run: { italics: true, color: t.colors.muted },
    },
  ];

  const characterStyles: readonly ICharacterStyleOptions[] = [
    {
      // The character-style counterpart of `Normal`, and the target of the
      // seven `w:basedOn w:val="DefaultParagraphFont"` references below. Like
      // `Normal` it declares nothing: `docDefaults` already has.
      id: STYLE_IDS.defaultParagraphFont,
      name: "Default Paragraph Font",
      semiHidden: true,
      unhideWhenUsed: true,
    },
    {
      id: STYLE_IDS.codeChar,
      name: "Code Char",
      basedOn: STYLE_IDS.defaultParagraphFont,
      uiPriority: 30,
      run: {
        font: mono,
        ...sized(t.sizes.code),
        color: t.colors.inlineCodeText,
        shading: { type: ShadingType.CLEAR, color: "auto", fill: t.colors.inlineCodeBackground },
      },
    },
    {
      id: STYLE_IDS.htmlChar,
      name: "Raw HTML Char",
      basedOn: STYLE_IDS.codeChar,
      uiPriority: 31,
      run: { font: mono, ...sized(t.sizes.code), color: t.colors.muted },
    },
    {
      id: STYLE_IDS.imagePlaceholderChar,
      name: "Image Placeholder Char",
      basedOn: STYLE_IDS.defaultParagraphFont,
      uiPriority: 36,
      run: { italics: true, color: t.colors.muted },
    },
  ];

  return {
    default: {
      // docDefaults: the root every other style inherits from.
      document: {
        run: { font: t.fonts.body, ...sized(t.sizes.body) },
        paragraph: {
          spacing: {
            after: t.spacing.paragraphAfter,
            line: t.spacing.line,
            lineRule: LineRuleType.AUTO,
          },
        },
      },
      title: {
        name: "Title",
        basedOn: STYLE_IDS.normal,
        next: STYLE_IDS.normal,
        quickFormat: true,
        uiPriority: 10,
        paragraph: {
          spacing: { before: 0, after: 240, line: t.spacing.line, lineRule: LineRuleType.AUTO },
        },
        run: { font: t.fonts.heading, ...sized(t.sizes.title), color: t.colors.heading },
      },
      heading1: headingStyle(t, 1),
      heading2: headingStyle(t, 2),
      heading3: headingStyle(t, 3),
      heading4: headingStyle(t, 4),
      heading5: headingStyle(t, 5),
      heading6: headingStyle(t, 6),
      // docx types `strong` as a paragraph style, so it gets no `basedOn`
      // here - a paragraph style based on a character style is meaningless.
      // Nothing references it: `**bold**` becomes direct <w:b/>, exactly as
      // Word itself writes it. It is declared only so the id exists.
      strong: { name: "Strong", run: { bold: true } },
      listParagraph: {
        name: "List Paragraph",
        basedOn: STYLE_IDS.normal,
        quickFormat: true,
        uiPriority: 34,
        paragraph: {
          // No `after` here on purpose, so the 160 twips from `docDefaults`
          // survive. `contextualSpacing` already suppresses it *between* items
          // (they share this style), which is exactly the tight look a list
          // wants - while the paragraph *after* the list is a different style,
          // so the gap that closes the list comes back. Word's own
          // `ListParagraph` is built the same way. Pinning `after: 0` here made
          // `contextualSpacing` redundant and welded every tight list to
          // whatever followed it.
          spacing: { line: t.spacing.line, lineRule: LineRuleType.AUTO },
          // Tight by default; loose lists override `after` per item.
          contextualSpacing: true,
        },
      },
      hyperlink: {
        name: "Hyperlink",
        basedOn: STYLE_IDS.defaultParagraphFont,
        run: { color: t.colors.link, underline: { type: UnderlineType.SINGLE } },
      },
      footnoteReference: {
        name: "footnote reference",
        basedOn: STYLE_IDS.defaultParagraphFont,
        semiHidden: true,
        unhideWhenUsed: true,
        run: { superScript: true },
      },
      footnoteText: {
        name: "footnote text",
        basedOn: STYLE_IDS.normal,
        link: "FootnoteTextChar",
        semiHidden: true,
        unhideWhenUsed: true,
        paragraph: { spacing: { after: 0, line: 240, lineRule: LineRuleType.AUTO } },
        run: sized(t.sizes.footnote),
      },
      footnoteTextChar: {
        name: "Footnote Text Char",
        basedOn: STYLE_IDS.defaultParagraphFont,
        link: "FootnoteText",
        semiHidden: true,
        run: sized(t.sizes.footnote),
      },
    },
    paragraphStyles,
    characterStyles,
  };
}
