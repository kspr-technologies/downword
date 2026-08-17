/**
 * Ambient declaration for `mermaid`.
 *
 * mermaid is an **optional** peer dependency: it is not installed in this
 * repository, and it must not be a hard dependency of a package whose whole
 * pitch is that it is small. Under `moduleResolution: "bundler"` that makes the
 * `import("mermaid")` in `engine.ts` an unresolved module, which `strict`
 * rejects — so the module is declared here rather than silenced with `any` or
 * `@ts-ignore`, both of which this repository bans.
 *
 * Only the two members `src/mermaid/**` calls are declared; see `engine.ts` for
 * why those shapes are owned here rather than borrowed from mermaid. The
 * declaration is structural, so a consumer who *does* have mermaid installed
 * gets their own types — nothing here reaches the published `.d.ts`.
 *
 * As in `src/highlight/vendor.d.ts`, the `import type` lives *inside* the
 * `declare module` body on purpose: a top-level import would turn this file
 * into a module, and `declare module` inside a module means *augmentation*,
 * which would fail against a package whose real types are not present.
 */
declare module "mermaid" {
  import type { MermaidApi } from "./engine.js";

  /**
   * mermaid's default export: the singleton API object.
   *
   * Typed as the two-method subset downword drives it through, which is a
   * strict subset of the real thing.
   */
  const mermaid: MermaidApi;
  export default mermaid;
}
