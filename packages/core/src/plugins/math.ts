import type { DownwordPlugin } from "../types.js";

/** Options for {@link mathPlugin}. */
export interface MathPluginOptions {
  /**
   * Which TeX delimiters to recognise.
   * `"dollars"` -> `$inline$` / `$$display$$`.
   */
  readonly delimiters?: "dollars";
  /** Render display equations centered. */
  readonly centerDisplay?: boolean;
}

/**
 * TeX -> OOXML math (OMML) support.
 *
 * TODO(T2): implement. The verified pipeline is
 *   `@vscode/markdown-it-katex` tokens (`math_inline` / `math_block`)
 *   -> `temml.renderToString(tex, { xml: true })` (MathML, DOM-free)
 *   -> `mathml2omml`'s named export `{ mml2omml }`
 *   -> unwrap `ImportedXmlComponent.fromXmlString(omml).root[0]`
 *      (`fromXmlString` cannot be inserted directly: it emits an `<undefined>`
 *      root tag and corrupts the package)
 *   -> cast to `ParagraphChild`.
 * The `Math*` builder classes in `docx` cannot express matrices, accents, bars
 * or equation arrays, so the OMML path is required rather than optional.
 *
 * Today this is a typed no-op so that the `downword/plugins/math` export is
 * resolvable and its option shape is locked in.
 */
export function mathPlugin(_options: MathPluginOptions = {}): DownwordPlugin {
  return {
    name: "math",
    setup: (context) => {
      context.warn("downword/plugins/math: not implemented yet; math will be emitted as raw text.");
    },
  };
}
