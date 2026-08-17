/**
 * Ambient declarations for markdown-it plugins that ship no types.
 *
 * `markdown-it-footnote@4` is a hard dependency of this package but publishes
 * plain JavaScript with no `.d.ts` and no DefinitelyTyped package. Declaring it
 * here (rather than reaching for `any` at the import site) keeps the
 * zero-`any`/zero-`@ts-ignore` rule intact.
 *
 * The `import type` lives *inside* the `declare module` body on purpose: a
 * top-level import would turn this file into a module, and `declare module`
 * inside a module means *augmentation*, which fails for a package that has no
 * types to augment.
 *
 * The plugin's real signature is `(md, options?) => void`; only the
 * zero-options form is used, so that is all that is declared.
 */
declare module "markdown-it-footnote" {
  import type { MarkdownIt } from "markdown-it";

  const footnotePlugin: (md: MarkdownIt) => void;
  export default footnotePlugin;
}
