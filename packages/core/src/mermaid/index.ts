/**
 * mermaid diagrams: the implementation.
 *
 * The public entry point is `@ksprtech/downword/plugins/mermaid` (`src/plugins/mermaid.ts`),
 * which re-exports this barrel. Everything here is browser-first — see
 * `types.ts` for why a DOM is not negotiable — and nothing in it is reachable
 * from the `@ksprtech/downword` entry, which `tests/bundle.test.ts` enforces.
 */

export {
  createBrowserMermaidRenderer,
  isEngineUnavailable,
  MERMAID_ENGINE_ERROR,
  unwrapMermaid,
  type BrowserRendererOptions,
  type MermaidApi,
  type MermaidConfig,
  type MermaidLoader,
  type MermaidModule,
  type MermaidRenderOutput,
} from "./engine.js";

export { detectDomSupport, type DomSupport } from "./environment.js";

export { isMermaidFence, readMermaidFence, type MermaidFenceInfo } from "./fence.js";

export {
  countMermaidDiagrams,
  renderMermaid,
  type MermaidOptions,
  type MermaidResult,
} from "./pass.js";

export { createCanvasRasterizer, type CanvasRasterizerOptions } from "./raster.js";

export {
  MERMAID_DIAGNOSTIC_CODES,
  mermaidWarningSeverity,
  type MermaidDiagnosticCode,
  type MermaidRenderRequest,
  type MermaidRenderer,
  type MermaidSvg,
  type MermaidWarning,
  type MermaidWarningHandler,
} from "./types.js";
