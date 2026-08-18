/**
 * `@ksprtech/downword/plugins/math` — TeX in, **native Word equations** out.
 *
 * ```ts
 * import { convertWithMath } from "@ksprtech/downword/plugins/math";
 *
 * const bytes = await convertWithMath("The area is $\\pi r^2$.");
 * ```
 *
 * Open the result in Word, click the equation, and it is a real equation: the
 * ribbon switches to Equation Tools, the cursor moves through the numerator and
 * the denominator, `\pi` is a symbol rather than a picture of one, and the
 * whole thing rescales with the paragraph and is found by search. Nothing here
 * produces an image unless you ask it to.
 *
 * ## What is in the box
 *
 * | export | job |
 * | --- | --- |
 * | {@link mathMarkdownIt} | `$…$` / `$$…$$` -> markdown-it tokens |
 * | {@link convertDocumentMath} | the async pass that fills each equation's `omml` slot |
 * | {@link mathPlugin} | both of the above, bundled, plus the plugin metadata |
 * | {@link convertWithMath} | `convert()` with those two steps wired in |
 *
 * The split exists because the pipeline is
 * `parse (sync) -> resolve (async) -> render (sync)`, and maths needs a foot in
 * the first two: a markdown-it rule to find the equations, and an `await` to
 * load the TeX engine. Use {@link convertWithMath} unless you are driving the
 * stages yourself.
 *
 * ## Why it is a separate entry point
 *
 * `temml` and `mathml2omml` are **optional peer dependencies**, reached through
 * a dynamic `import()` in `src/math/engine.ts`. Importing `@ksprtech/downword` therefore
 * costs nothing extra, and `tests/bundle.test.ts` proves it by bundling the
 * main entry and failing if either package appears in the graph. Install them
 * alongside this entry:
 *
 * ```sh
 * npm i @ksprtech/downword temml mathml2omml
 * ```
 *
 * Without them every equation degrades to its TeX source as literal text, with
 * one `engine-unavailable` warning saying so. That is the rule everywhere in
 * this plugin: **nothing throws, and no equation is silently dropped.**
 *
 * @module
 */

import { Packer, type Document } from "docx";

import { DownwordError, wrapError } from "../errors.js";
import { createImageResolver, resolveDocumentImages } from "../images/index.js";
import { convertDocumentMath } from "../math/document.js";
import { mathMarkdownIt } from "../math/markdown-it.js";
import { resolveMathOptions, type MathOptions } from "../math/options.js";
import type { DocumentNode } from "../model.js";
import {
  resolveConvertOptions,
  type ConvertOptions,
  type ConvertWarning,
  type ConvertWarningHandler,
} from "../options.js";
import { parseMarkdown, type MarkdownItPlugin } from "../parse/index.js";
import { prepareHighlights, renderDocument, type HighlightMap } from "../render/index.js";
import type { DownwordPlugin } from "./types.js";

export { createMathConverter, type MathConverter, type MathResult } from "../math/convert.js";
export { convertDocumentMath, type MathDocumentConversion } from "../math/document.js";
export {
  captureConsoleWarnings,
  loadMathEngine,
  type MathEngine,
  type MathEngineLoader,
  type TexRenderOptions,
} from "../math/engine.js";
export {
  estimateMathBox,
  mathImageNode,
  mathSource,
  rasterizeMath,
  wrapMathmlInSvg,
} from "../math/image.js";
export { mathMarkdownIt, type MathMarkdownItOptions } from "../math/markdown-it.js";
export { checkOmml, OMML_NAMESPACE, WML_NAMESPACE, type OmmlCheck } from "../math/omml.js";
export {
  DEFAULT_FONT_SIZE,
  DEFAULT_MAX_LENGTH,
  resolveMathOptions,
  type MathMode,
  type MathOptions,
  type ResolvedMathOptions,
} from "../math/options.js";
export {
  MATH_WARNING_CODES,
  mathWarningSeverity,
  type MathWarning,
  type MathWarningCode,
  type MathWarningHandler,
} from "../math/warnings.js";

/* -------------------------------------------------------------------------- */
/* The plugin object                                                           */
/* -------------------------------------------------------------------------- */

/**
 * The math plugin: a {@link DownwordPlugin} with both halves of the pipeline
 * hanging off it.
 *
 * `DownwordPlugin` describes what every `@ksprtech/downword/plugins/*` entry has in
 * common — a name, and a chance to warn during setup. Maths needs more than
 * that (a markdown-it rule **and** an async pass), so it adds two members
 * rather than pretending to fit.
 */
export interface MathDownwordPlugin extends DownwordPlugin {
  readonly name: "math";
  /** Give this to `ConvertOptions.plugins`. */
  readonly markdownIt: MarkdownItPlugin;
  /** Run this between `parseMarkdown` and `renderDocument`. */
  readonly convertDocument: (document: DocumentNode) => ReturnType<typeof convertDocumentMath>;
}

/**
 * Builds the math plugin with its options baked in.
 *
 * ```ts
 * const math = mathPlugin({ macros: { "\\RR": "\\mathbb{R}" } });
 *
 * const doc = parseMarkdown(source, { plugins: [math.markdownIt] });
 * const { document } = await math.convertDocument(doc);
 * const file = renderDocument(document);
 * ```
 *
 * @param options - See {@link MathOptions}.
 * @returns The plugin.
 * @throws {DownwordError} `"invalid-options"` — validation happens here, once,
 *   rather than on every document.
 */
export function mathPlugin(options: MathOptions = {}): MathDownwordPlugin {
  // Fail immediately on a bad option, exactly as `convert()` does, instead of
  // at the first equation.
  const resolved = resolveMathOptions(options);

  return {
    name: "math",
    markdownIt: mathMarkdownIt({ delimiters: resolved.delimiters }),
    convertDocument: (document) => convertDocumentMath(document, options),
  };
}

/* -------------------------------------------------------------------------- */
/* The one-call form                                                           */
/* -------------------------------------------------------------------------- */

/** {@link ConvertOptions} plus the maths. */
export interface ConvertWithMathOptions extends ConvertOptions {
  /**
   * How equations are handled. See {@link MathOptions}.
   *
   * Math warnings arrive on `math.onWarning`, not on
   * {@link ConvertOptions.onWarning}: they carry the TeX that failed, which the
   * three-stage `ConvertWarning` union has nowhere to put.
   */
  readonly math?: MathOptions | undefined;
}

/** Warning handler that drops everything. */
const IGNORE: ConvertWarningHandler = () => {};

/** Wraps a host handler so a logger that throws cannot fail a conversion. */
function createEmitter(handler: ConvertWarningHandler | undefined): ConvertWarningHandler {
  if (handler === undefined) return IGNORE;
  return (warning: ConvertWarning) => {
    try {
      handler(warning);
    } catch {
      /* losing a warning beats losing the document */
    }
  };
}

/**
 * `convertToDocument()`, with the equations converted.
 *
 * Exactly the core pipeline — same options, same warnings, same defaults — plus
 * {@link mathMarkdownIt} appended to the parser's plugins and
 * {@link convertDocumentMath} run between the parse and the render. It is built
 * out of `downword`'s own public functions rather than forking `convert()`, and
 * `tests/math.test.ts` asserts that a document containing no maths comes out of
 * it byte-identical to `convert()`'s.
 *
 * @param markdown - The source.
 * @param options - See {@link ConvertWithMathOptions}.
 * @returns The `docx` `Document`, one step short of the zip.
 * @throws {DownwordError} and nothing else.
 */
export async function convertToDocumentWithMath(
  markdown: string,
  options: ConvertWithMathOptions = {},
): Promise<Document> {
  if (typeof markdown !== "string") {
    throw new DownwordError(
      "invalid-input",
      `markdown must be a string, got ${markdown === null ? "null" : typeof markdown}`,
    );
  }

  const emit = createEmitter(options.onWarning);
  const resolved = resolveConvertOptions(options, emit);
  const math = options.math ?? {};
  // Validated up front, so a typo in `math.math` fails before any parsing.
  const resolvedMath = resolveMathOptions(math);

  // `math: "off"` still needs the tokens: the renderer's fallback prints the
  // TeX source, which is exactly what that mode promises.
  const plugins: MarkdownItPlugin[] = [
    ...(resolved.parse.plugins ?? []),
    mathMarkdownIt({ delimiters: resolvedMath.delimiters }),
  ];

  let document: DocumentNode;
  try {
    document = parseMarkdown(markdown, { ...resolved.parse, plugins });
  } catch (cause: unknown) {
    throw wrapError("parse-failed", "could not parse the markdown", cause);
  }

  const converted = await convertDocumentMath(document, math);

  // From here on this is `convertToDocument()` verbatim, against the rewritten
  // document: images and highlighting are independent and both I/O-bound.
  const resolver =
    resolved.imageResolver ?? createImageResolver({ allowRemote: resolved.allowRemoteImages });
  const [{ images }, highlights] = await Promise.all([
    resolveDocumentImages(converted.document, resolver, {
      onDiagnostic: (diagnostic) => {
        emit({ ...diagnostic, stage: "image" });
      },
    }),
    resolved.highlighter === null
      ? Promise.resolve<HighlightMap>(new Map())
      : prepareHighlights(converted.document, resolved.highlighter, {
          onWarning: (warning) => {
            emit({ ...warning, stage: "render" });
          },
        }),
  ]);

  try {
    return renderDocument(converted.document, { ...resolved.render, images, highlights });
  } catch (cause: unknown) {
    throw wrapError("render-failed", "could not render the document", cause);
  }
}

/**
 * `convert()`, with the equations converted. See {@link convertToDocumentWithMath}.
 *
 * @param markdown - The source.
 * @param options - See {@link ConvertWithMathOptions}.
 * @returns The `.docx` bytes.
 * @throws {DownwordError} and nothing else.
 */
export async function convertWithMath(
  markdown: string,
  options: ConvertWithMathOptions = {},
): Promise<Uint8Array> {
  const file = await convertToDocumentWithMath(markdown, options);
  try {
    return new Uint8Array(await Packer.toArrayBuffer(file));
  } catch (cause: unknown) {
    throw wrapError("pack-failed", "could not build the .docx package", cause);
  }
}
