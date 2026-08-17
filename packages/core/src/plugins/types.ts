/**
 * The plugin shape the `downword/plugins/*` entry points share.
 *
 * Deliberately **not** re-exported from `downword`'s main entry, and
 * deliberately not the same thing as `ConvertOptions.plugins` — which takes
 * plain `MarkdownItPlugin`s, the extension point that works today.
 *
 * The distinction is not bureaucracy. Math and mermaid are *two-sided*: math
 * needs a markdown-it plugin to produce `math_inline`/`math_block` tokens **and**
 * an async pass to fill each node's `omml` slot (`temml` → `mathml2omml`);
 * mermaid needs to recognise a fence **and** rasterise an SVG so `ImageRun` has
 * the raster twin OOXML requires. Neither fits in a `(md) => void`. This type
 * is the seam those passes will hang off, kept stable from the first release so
 * the published `exports` map does not have to change when they land.
 *
 * @see `src/plugins/math.ts`, `src/plugins/mermaid.ts`
 */

/**
 * A downword plugin.
 *
 * Not yet accepted by `convert()`: both implementations are stubs. See
 * `ConvertOptions.plugins` for the markdown-it extension point that is live.
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
  /** Collects a non-fatal warning. */
  readonly warn: (message: string) => void;
}
