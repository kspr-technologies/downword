/**
 * downword - Markdown to Word (.docx), entirely in your browser.
 *
 * Public entry point. Subpath entry points:
 *  - `downword/plugins/math`    -> src/plugins/math.ts
 *  - `downword/plugins/mermaid` -> src/plugins/mermaid.ts
 */

export { convert, VERSION } from "./convert.js";
export { DOCX_MIME_TYPE } from "./types.js";
export type {
  ConvertOptions,
  ConvertResult,
  DownwordPlugin,
  DownwordPluginContext,
} from "./types.js";
