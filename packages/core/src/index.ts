/**
 * downword — Markdown to Word (.docx), entirely in your browser.
 *
 * ```ts
 * import { convert } from "@ksprtech/downword";
 *
 * const bytes = await convert("# Hello\n\nFrom **downword**.");
 * ```
 *
 * ## What is in this entry point
 *
 * Everything needed to convert markdown, plus both halves of the pipeline
 * (`parseMarkdown` and `renderDocument`) for callers who want to inspect the
 * document in between. It is free of Node builtins and of the DOM, so it runs
 * unchanged in a tab, a worker, Node and a server handler — a property
 * `tests/bundle.test.ts` enforces by bundling this file for the browser and
 * inspecting the import graph.
 *
 * ## What is deliberately not
 *
 * | subpath | why it is separate |
 * | --- | --- |
 * | `@ksprtech/downword/highlight` | highlight.js is an **optional peer dependency**; importing it here would put 62 lazily-loaded grammar chunks into every consumer's build. |
 * | `@ksprtech/downword/images/node` | the only module that touches `node:fs`/`node:dns`. Keeping it out means a browser bundler never has to resolve them. |
 * | `@ksprtech/downword/plugins/math` | `temml` and `mathml2omml` are **optional peer dependencies**, dynamically imported; the entry also carries the async pass that fills each equation's `omml` slot. |
 * | `@ksprtech/downword/plugins/mermaid` | mermaid is a ~500 kB **optional peer dependency**, dynamically imported, and the pass needs a DOM. |
 *
 * @module
 */

/* -------------------------------------------------------------------------- */
/* The API                                                                     */
/* -------------------------------------------------------------------------- */

export { convert, convertToBlob, convertToDocument, DOCX_MIME_TYPE, VERSION } from "./convert.js";

export { DownwordError, isDownwordError, type DownwordErrorCode } from "./errors.js";

export {
  resolveConvertOptions,
  THEMES,
  type ConvertMetadata,
  type ConvertOptions,
  type ConvertWarning,
  type ConvertWarningHandler,
  type ConvertWarningStage,
  type HtmlHandling,
  type LineBreakHandling,
  type MarginsInit,
  type ResolvedConvertOptions,
  type ThemeName,
} from "./options.js";

/* -------------------------------------------------------------------------- */
/* Stage 1: markdown -> model                                                  */
/* -------------------------------------------------------------------------- */

export {
  createMarkdownIt,
  MAX_BLOCK_NESTING,
  PARSE_WARNING_CODES,
  parseMarkdown,
  parseTokens,
  parseWarningSeverity,
  resolveParseOptions,
  type MarkdownItPlugin,
  type ParseOptions,
  type ParseWarning,
  type ParseWarningCode,
  type ParseWarningHandler,
  type ResolvedParseOptions,
} from "./parse/index.js";

/* -------------------------------------------------------------------------- */
/* Stage 2: model -> docx                                                      */
/* -------------------------------------------------------------------------- */

export {
  buildStyles,
  DEFAULT_THEME,
  headingStyleId,
  prepareHighlights,
  prepareImages,
  RENDER_WARNING_CODES,
  renderDocument,
  renderWarningSeverity,
  resolveTheme,
  scopeColor,
  STYLE_IDS,
  type BaseWarning,
  type BookmarkAnchor,
  type BulletLevel,
  type FootnoteIndex,
  type FootnotePolicy,
  type HeadingSpec,
  type HighlightColorValue,
  type Highlighter,
  type HighlightMap,
  type HighlightSpan,
  type HtmlPolicy,
  type ImageMap,
  type ImageResolver,
  type LevelFormatValue,
  type PageInit,
  type PageNumberFormat,
  type PageNumberSettings,
  type PageNumbersInit,
  type PageSettings,
  type PageSize,
  type PrepareOptions,
  type RenderOptions,
  type RenderWarning,
  type RenderWarningCode,
  type RenderWarningHandler,
  type ResolvedRenderOptions,
  type SoftBreakPolicy,
  type TaskGlyphs,
  type TextDirection,
  type TocInit,
  type TocSettings,
  type Theme,
  type ThemeColors,
  type ThemeFonts,
  type ThemeInit,
  type ThemeSizes,
  type ThemeSpacing,
  type WarningSeverity,
} from "./render/index.js";

/* -------------------------------------------------------------------------- */
/* Images                                                                      */
/* -------------------------------------------------------------------------- */

export {
  createImageResolver,
  hasImageErrors,
  isObservableImageResolver,
  NULL_IMAGE_RESOLVER,
  resolveDocumentImages,
  type DocumentImageResolution,
  type ImageDiagnostic,
  type ImageDiagnosticCode,
  type ImageDiagnosticHandler,
  type ImageDiagnosticSeverity,
  type ImageFetchPolicy,
  type ImageLoader,
  type ImageRasterizer,
  type ImageResolverOptions,
  type ObservableImageResolver,
  type ResolveDocumentImagesOptions,
} from "./images/index.js";

/* -------------------------------------------------------------------------- */
/* The document model                                                          */
/* -------------------------------------------------------------------------- */

export {
  childrenOf,
  findMark,
  hasMark,
  isBlockNode,
  isInlineNode,
  isStructuralNode,
  nodesOfType,
  nodeText,
  slugify,
  visit,
  type BlockNode,
  type BlockquoteNode,
  type CodeBlockNode,
  type CodeHighlightSpan,
  type DocumentMetadata,
  type DocumentNode,
  type FootnoteDefinitionNode,
  type FootnoteReferenceNode,
  type HardBreakNode,
  type HeadingLevel,
  type HeadingNode,
  type HtmlBlockNode,
  type HtmlInlineNode,
  type ImageFormat,
  type ImageNode,
  type InlineNode,
  type LinkMark,
  type ListItemNode,
  type ListKind,
  type ListNode,
  type Mark,
  type MarkOfType,
  type MarkType,
  type MathBlockNode,
  type MathInlineNode,
  type Node,
  type NodeCategory,
  type NodeOfType,
  type NodeType,
  type ParagraphNode,
  type RasterImageFormat,
  type ResolvedImage,
  type ResolvedRasterImage,
  type SoftBreakNode,
  type StructuralNode,
  type TableAlignment,
  type TableCellNode,
  type TableNode,
  type TableRowNode,
  type TextNode,
  type ThematicBreakNode,
  type VisitAction,
  type VisitContext,
  type VisitEnter,
  type VisitLeave,
} from "./model.js";

/* -------------------------------------------------------------------------- */
/* Units                                                                       */
/* -------------------------------------------------------------------------- */

export { inchesToTwips, pointsToTwips, TWIPS_PER_INCH, TWIPS_PER_POINT } from "./render/units.js";
