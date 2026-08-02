import type { DownwordPlugin } from "../types.js";

/** Options for {@link mermaidPlugin}. */
export interface MermaidPluginOptions {
  /** Diagram width in pixels (docx converts px -> EMU with a x9525 factor). */
  readonly width?: number;
  /** Diagram height in pixels. */
  readonly height?: number;
  /** Emit the SVG plus a raster fallback (required by `ImageRun` for SVG). */
  readonly rasterFallback?: boolean;
}

/**
 * ```mermaid``` fenced blocks -> embedded diagrams.
 *
 * TODO(T4): implement. Constraints already established:
 *  - `ImageRun` requires an explicit `type` in docx v9 (no sniffing), and
 *    `type: "svg"` requires a `fallback` raster image.
 *  - String `data` is interpreted as **base64 only**; raw SVG markup throws
 *    `InvalidCharacterError` from `atob`. Pass
 *    `new TextEncoder().encode(svg)` instead.
 *  - `ImageRun` silently drops `solidFill`.
 *
 * Today this is a typed no-op so that the `downword/plugins/mermaid` export is
 * resolvable and its option shape is locked in.
 */
export function mermaidPlugin(_options: MermaidPluginOptions = {}): DownwordPlugin {
  return {
    name: "mermaid",
    setup: (context) => {
      context.warn(
        "downword/plugins/mermaid: not implemented yet; mermaid blocks stay as code blocks.",
      );
    },
  };
}
