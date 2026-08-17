/**
 * Ambient declaration for `highlight.js/lib/core`.
 *
 * `highlight.js@11` publishes types for the package root and — usefully — for
 * the `highlight.js/lib/languages/*` subpath, which it declares as a wildcard
 * ambient module (`types/index.d.ts`, `const defineLanguage: LanguageFn`). So
 * the 62 grammar imports in `languages.ts` are already typed.
 *
 * `./lib/core` is the gap: its entry in the package's `exports` map carries
 * only `require`/`import`, no `types` condition, and no ambient declaration
 * covers it. Under `moduleResolution: "bundler"` that makes
 * `import("highlight.js/lib/core")` an untyped import, which `strict` rejects —
 * so it is declared here rather than silenced with `any` or `@ts-ignore`.
 *
 * As in `src/parse/vendor.d.ts`, the `import type` lives *inside* the
 * `declare module` body on purpose: a top-level import would turn this file
 * into a module, and `declare module` inside a module means *augmentation*,
 * which fails for a subpath that has no types to augment.
 *
 * Only the members `src/highlight/**` calls are declared; see `engine.ts` for
 * why those shapes are owned here rather than borrowed from highlight.js.
 */
declare module "highlight.js/lib/core" {
  import type { HighlightEngine } from "./engine.js";

  /**
   * The bare highlight.js core: the parser and the registry, with **no**
   * grammars registered. Everything must be added with `registerLanguage`,
   * which is exactly what makes per-language lazy loading possible.
   */
  const hljs: HighlightEngine;
  export default hljs;
}
