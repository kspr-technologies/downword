/**
 * Options for the math pass, and their defaults.
 *
 * Validation goes through `src/validate.ts` so that a bad value fails with the
 * same `DownwordError` code, in the same words, as a bad value anywhere else in
 * the package.
 */

import type { ImageRasterizer } from "../images/types.js";
import { integer, invalid, oneOf } from "../validate.js";
import type { MathEngineLoader } from "./engine.js";
import { loadMathEngine } from "./engine.js";
import type { MathWarningHandler } from "./warnings.js";

/**
 * What to do with each equation.
 *
 * - `"omml"` — a native Word equation (`<m:oMath>`). Editable in Word's
 *   equation editor, searchable, and it reflows and rescales with the document.
 *   **The default**, and the reason this plugin exists.
 * - `"image"` — a picture. Pixel-perfect and immune to Word's MathML gaps, but
 *   dead: not editable, not searchable, and blurry when scaled. Needs a
 *   host-supplied {@link MathOptions.rasterizer}; see {@link MathOptions.math}.
 * - `"off"` — no conversion at all. The renderer's existing fallback puts the
 *   TeX source in the document as literal monospace text, and neither engine is
 *   loaded.
 */
export type MathMode = "omml" | "image" | "off";

/** Longest equation the converter will attempt, in UTF-16 code units. */
export const DEFAULT_MAX_LENGTH = 200_000;

/** Nominal font size the `"image"` mode draws at, in CSS pixels. */
export const DEFAULT_FONT_SIZE = 16;

/** Options for {@link import("./document.js").convertDocumentMath} and the plugin. */
export interface MathOptions {
  /**
   * Which representation each equation gets. Defaults to `"omml"`.
   *
   * `"image"` needs a {@link MathOptions.rasterizer}, because turning MathML
   * into pixels needs a layout engine and this package does not contain one —
   * the same position `downword` already takes for SVG images. In Node, with no
   * rasteriser supplied, `"image"` reports `image-no-rasterizer` and falls back
   * to OMML rather than dropping the equation; it does not pretend to have
   * rendered a picture.
   */
  readonly math?: MathMode | undefined;
  /** Which delimiters {@link import("./markdown-it.js").mathMarkdownIt} recognises. Only `"dollars"` today. */
  readonly delimiters?: "dollars" | undefined;
  /**
   * Longest equation to convert, in UTF-16 code units. Defaults to 200 000.
   *
   * A backstop against a pathological paste, not a normal limit: `temml` turns
   * a 100 kB equation into ~1 MB of MathML in under 100 ms, so the default has
   * an order of magnitude of headroom. Above it the equation stays literal text
   * and a `tex-too-large` warning is raised.
   */
  readonly maxLength?: number | undefined;
  /**
   * TeX macros, e.g. `{ "\\RR": "\\mathbb{R}" }`.
   *
   * Bodies must be strings; `temml`'s function-valued macros are deliberately
   * not accepted, because a document should not be able to run code. Each
   * equation gets a fresh copy, so a `\newcommand` in one equation cannot leak
   * into the next.
   */
  readonly macros?: Readonly<Record<string, string>> | undefined;
  /**
   * Turns the equation's SVG into the raster twin OOXML requires. Only used by
   * `math: "image"`.
   *
   * The same {@link ImageRasterizer} interface `downword` already uses for SVG
   * pictures, so a host that has one has one for both.
   */
  readonly rasterizer?: ImageRasterizer | undefined;
  /** Nominal font size for `math: "image"`, in CSS pixels. Defaults to 16. */
  readonly fontSize?: number | undefined;
  /** Supplies `temml` and `mathml2omml`. Defaults to importing them. */
  readonly load?: MathEngineLoader | undefined;
  /** Called once per non-fatal problem. */
  readonly onWarning?: MathWarningHandler | undefined;
}

/** A {@link MathOptions} with every default filled in. */
export interface ResolvedMathOptions {
  readonly math: MathMode;
  readonly delimiters: "dollars";
  readonly maxLength: number;
  readonly macros: Readonly<Record<string, string>>;
  readonly rasterizer: ImageRasterizer | null;
  readonly fontSize: number;
  readonly load: MathEngineLoader;
  readonly onWarning: MathWarningHandler | null;
}

/** No macros. Shared, and never handed out without being copied first. */
const NO_MACROS: Readonly<Record<string, string>> = Object.freeze({});

function checkMacros(macros: MathOptions["macros"]): Readonly<Record<string, string>> {
  if (macros === undefined) return NO_MACROS;
  if (typeof macros !== "object" || macros === null || Array.isArray(macros)) {
    throw invalid(`options.math.macros must be an object, got ${JSON.stringify(macros)}`);
  }
  for (const [name, body] of Object.entries(macros)) {
    if (typeof body !== "string") {
      throw invalid(
        `options.math.macros[${JSON.stringify(name)}] must be a TeX string, got ${typeof body}`,
      );
    }
  }
  return macros;
}

function checkRasterizer(rasterizer: MathOptions["rasterizer"]): ImageRasterizer | null {
  if (rasterizer === undefined) return null;
  if (
    typeof rasterizer !== "object" ||
    rasterizer === null ||
    typeof rasterizer.rasterize !== "function"
  ) {
    throw invalid("options.math.rasterizer must be an ImageRasterizer with a rasterize() method");
  }
  return rasterizer;
}

function checkLoader(load: MathOptions["load"]): MathEngineLoader {
  if (load === undefined) return loadMathEngine;
  if (typeof load !== "function") {
    throw invalid("options.math.load must be a function returning a Promise<MathEngine>");
  }
  return load;
}

function checkHandler(handler: MathOptions["onWarning"]): MathWarningHandler | null {
  if (handler === undefined) return null;
  if (typeof handler !== "function") {
    throw invalid("options.math.onWarning must be a function");
  }
  return handler;
}

/**
 * Fills in every default and rejects values the pass cannot use.
 *
 * @param options - Whatever the caller passed.
 * @returns The resolved options.
 * @throws {import("../errors.js").DownwordError} `"invalid-options"`.
 */
export function resolveMathOptions(options: MathOptions = {}): ResolvedMathOptions {
  if (typeof options !== "object" || options === null) {
    throw invalid(`options.math must be an object, got ${JSON.stringify(options)}`);
  }

  return {
    math: oneOf(options.math, ["omml", "image", "off"] as const, "math.math", "omml"),
    delimiters: oneOf(options.delimiters, ["dollars"] as const, "math.delimiters", "dollars"),
    maxLength: integer(options.maxLength, "math.maxLength", DEFAULT_MAX_LENGTH, 1, 2 ** 31 - 1),
    macros: checkMacros(options.macros),
    rasterizer: checkRasterizer(options.rasterizer),
    fontSize: integer(options.fontSize, "math.fontSize", DEFAULT_FONT_SIZE, 4, 400),
    load: checkLoader(options.load),
    onWarning: checkHandler(options.onWarning),
  };
}
