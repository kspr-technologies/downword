/**
 * Public types for downword.
 *
 * Everything in this file is part of the package's public API surface and is
 * re-exported from `src/index.ts`. Breaking changes here need a changeset.
 */

/** Options accepted by {@link convert}. */
export interface ConvertOptions {
  /** Value written to the document's `dc:title` core property. */
  readonly title?: string;
  /** Value written to the document's `dc:creator` core property. */
  readonly creator?: string;
  /** Value written to the document's `dc:description` core property. */
  readonly description?: string;
  /**
   * Plugins that extend the (not yet implemented) renderer.
   * See `src/plugins/math.ts` and `src/plugins/mermaid.ts`.
   */
  readonly plugins?: readonly DownwordPlugin[];
}

/** Result of a successful conversion. */
export interface ConvertResult {
  /** The raw `.docx` bytes. Write them to disk or wrap them in a `Blob`. */
  readonly bytes: Uint8Array;
  /** Always {@link DOCX_MIME_TYPE}; handy when building a `Blob`/`Response`. */
  readonly mimeType: string;
  /** Non-fatal problems encountered while converting. Never `undefined`. */
  readonly warnings: readonly string[];
}

/**
 * A downword plugin.
 *
 * The renderer that consumes these does not exist yet (T0 scaffold); the shape
 * is fixed now so that the published `exports` map and the plugin entry points
 * are stable from the first release.
 */
export interface DownwordPlugin {
  /** Unique, kebab-case plugin name, e.g. `"math"`. */
  readonly name: string;
  /** Called once per conversion, before any markdown is parsed. */
  readonly setup?: (context: DownwordPluginContext) => void;
}

/** Handed to {@link DownwordPlugin.setup}. */
export interface DownwordPluginContext {
  /** Version of downword running the plugin. */
  readonly version: string;
  /** Collects a non-fatal warning that ends up in {@link ConvertResult.warnings}. */
  readonly warn: (message: string) => void;
}

/** MIME type of an OOXML WordprocessingML document. */
export const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
