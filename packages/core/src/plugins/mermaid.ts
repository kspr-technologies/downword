/**
 * `@ksprtech/downword/plugins/mermaid` — ```` ```mermaid ```` fences become embedded diagrams.
 *
 * ```ts
 * import { parseMarkdown, renderDocument } from "@ksprtech/downword";
 * import { renderMermaid } from "@ksprtech/downword/plugins/mermaid";
 *
 * const { document, warnings } = await renderMermaid(parseMarkdown(markdown));
 * const file = renderDocument(document);
 * ```
 *
 * ## Three things to know before using it
 *
 * **1. It needs a browser.** mermaid measures text by laying it out in a DOM,
 * so there is no headless path that is not a headless *browser*. In Node the
 * pass leaves every fence exactly as it found it — the diagram's source is
 * still in the document, as a code block — and reports one `no-dom` warning
 * saying so. Nothing throws, and nothing disappears. To render in Node, supply
 * both seams yourself: a `renderer` (mermaid driven through jsdom, or a
 * headless browser) and a `rasterizer` (`resvg`, `sharp`, `@napi-rs/canvas`).
 *
 * **2. mermaid is an optional peer dependency.** It is ~500 kB, and it is
 * loaded through `await import("mermaid")` on the first diagram — so a document
 * with none fetches nothing, and a project that never imports this subpath does
 * not pay for it at all. `tests/bundle.test.ts` bundles the main entry and
 * fails if mermaid appears anywhere in its import graph.
 *
 * **3. What lands in the file is a picture, not a vector.** OOXML cannot hold a
 * bare SVG: it stores a raster *plus* an `asvg:svgBlip` extension, and `docx`'s
 * `ImageRun` requires the raster twin for `type: "svg"`. The default is
 * therefore a PNG rendered at twice its display size — sharp on screen and in
 * print, and drawn by every reader. `embed: "svg"` ships both halves, which
 * keeps the diagram crisp at any zoom in Word 2016+ at the cost of carrying two
 * copies.
 *
 * ## What it produces
 *
 * Each fence becomes a centred `Figure` paragraph holding the image, with alt
 * text from the caption (or a description of the diagram kind, e.g. "Mermaid
 * sequence diagram"), followed by an italic caption paragraph when the fence
 * carried one:
 *
 * ````md
 * ```mermaid The request pipeline
 * flowchart LR
 *   Browser --> Worker --> Word
 * ```
 * ````
 *
 * A caption can also come from mermaid's own `title:` frontmatter. Without one,
 * only the picture is emitted.
 *
 * @module
 */

export {
  countMermaidDiagrams,
  createBrowserMermaidRenderer,
  createCanvasRasterizer,
  detectDomSupport,
  isEngineUnavailable,
  isMermaidFence,
  MERMAID_DIAGNOSTIC_CODES,
  MERMAID_ENGINE_ERROR,
  mermaidWarningSeverity,
  readMermaidFence,
  renderMermaid,
  unwrapMermaid,
  type BrowserRendererOptions,
  type CanvasRasterizerOptions,
  type DomSupport,
  type MermaidApi,
  type MermaidConfig,
  type MermaidDiagnosticCode,
  type MermaidFenceInfo,
  type MermaidLoader,
  type MermaidModule,
  type MermaidOptions,
  type MermaidRenderOutput,
  type MermaidRenderRequest,
  type MermaidRenderer,
  type MermaidResult,
  type MermaidSvg,
  type MermaidWarning,
  type MermaidWarningHandler,
} from "../mermaid/index.js";
