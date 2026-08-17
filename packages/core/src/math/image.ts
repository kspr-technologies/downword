/**
 * `math: "image"` — the equation as a picture.
 *
 * ## Why this needs the host's help
 *
 * OOXML stores an SVG as the SVG **plus a raster twin** (`asvg:svgBlip` next to
 * `a:blip`), and `docx`'s `ImageRun` makes the twin mandatory, so "just embed
 * the SVG" is not an option — see `src/images/decode.ts`. Producing that twin
 * means laying out and painting MathML, which needs a browser (or a headless
 * one, or `resvg`). `downword` does not contain a layout engine and does not
 * want the dependency, which is exactly the position it already takes for SVG
 * pictures: rasterising is injected, through {@link ImageRasterizer}.
 *
 * So `math: "image"` is **browser-shaped**: it works when the host passes a
 * rasteriser backed by something that can draw. In Node, with nothing supplied,
 * it says so (`image-no-rasterizer`) and falls back to OMML — which is the
 * better representation anyway — instead of pretending or dropping the equation.
 *
 * ## Sizing
 *
 * The equation's height is estimated from the MathML's *structure*, which is
 * the part that estimates well: one line of inline math is about 1.4em, each
 * extra row of an `mtable` adds another, and a fraction or a stacked limit adds
 * roughly two thirds of one. The **width then comes from the rasteriser's own
 * aspect ratio**, because a rasteriser had to lay the equation out to draw it
 * and therefore knows the shape exactly. The result is a picture that is never
 * distorted and is at worst uniformly a little large or small.
 */

import { image, type ImageNode, type ResolvedImage, type ResolvedRasterImage } from "../model.js";
import { isValidResolvedImage } from "../images/decode.js";
import type { ImageRasterizer } from "../images/types.js";
import type { ResolvedMathOptions } from "./options.js";
import { abbreviateTex, type MathWarningSink } from "./warnings.js";

/** Widest picture we will claim, in CSS pixels. The renderer clamps again. */
const MAX_WIDTH = 4000;

/** Encodes markup as the bytes `ImageRun` needs (a string would be read as base64). */
const encode = (markup: string): Uint8Array => new TextEncoder().encode(markup);

/**
 * Wraps MathML in an SVG a rasteriser can draw.
 *
 * `foreignObject` is the only way to put MathML inside SVG, and it is the
 * reason this path is browser-shaped: a browser draws it, while `resvg` and
 * `sharp` ignore it. A rasteriser that cannot is expected to return `null`,
 * which degrades this equation to OMML.
 */
export function wrapMathmlInSvg(mathml: string, width: number, height: number): string {
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" ` +
    `viewBox="0 0 ${width} ${height}">` +
    `<foreignObject x="0" y="0" width="${width}" height="${height}">` +
    mathml +
    `</foreignObject></svg>`
  );
}

/** Counts non-overlapping occurrences of a literal needle. Linear, no regex. */
function count(haystack: string, needle: string): number {
  let total = 0;
  let at = haystack.indexOf(needle);
  while (at !== -1) {
    total += 1;
    at = haystack.indexOf(needle, at + needle.length);
  }
  return total;
}

/** The visible characters of a MathML fragment, tags removed. */
function textLength(mathml: string): number {
  let length = 0;
  let index = 0;
  while (index < mathml.length) {
    const lt = mathml.indexOf("<", index);
    const end = lt === -1 ? mathml.length : lt;
    length += mathml.slice(index, end).trim().length;
    if (lt === -1) break;
    const gt = mathml.indexOf(">", lt);
    if (gt === -1) break;
    index = gt + 1;
  }
  return length;
}

/** A first guess at the box an equation occupies, in CSS pixels. */
export function estimateMathBox(
  mathml: string,
  fontSize: number,
): { readonly width: number; readonly height: number } {
  const rows = Math.max(1, count(mathml, "<mtr"));
  // A fraction, a radical or a stacked limit each make the line taller.
  const stacked = Math.min(
    3,
    count(mathml, "<mfrac") + count(mathml, "<munderover") + count(mathml, "<msqrt"),
  );
  const height = Math.round(fontSize * (1.4 * rows + 0.7 * stacked));
  // 0.6em per visible character is the usual rule of thumb for a mixed
  // italic/upright maths line; the rasteriser corrects the shape afterwards.
  const width = Math.round(
    Math.min(MAX_WIDTH, Math.max(fontSize, textLength(mathml) * fontSize * 0.6)),
  );
  return { width, height };
}

/**
 * Draws one equation, or returns `null` and says why.
 *
 * @param mathml - `temml`'s output for this equation.
 * @param tex - The source, for diagnostics and alt text.
 * @param display - `$$…$$` rather than `$…$`.
 * @param options - Resolved math options; supplies the rasteriser and font size.
 * @param sink - Where warnings go.
 * @returns A {@link ResolvedImage} carrying the SVG and its raster twin, or
 *   `null`. Never throws.
 */
export async function rasterizeMath(
  mathml: string,
  tex: string,
  display: boolean,
  options: ResolvedMathOptions,
  sink: MathWarningSink,
): Promise<ResolvedImage | null> {
  const rasterizer: ImageRasterizer | null = options.rasterizer;
  if (rasterizer === null) return null;

  const box = estimateMathBox(mathml, options.fontSize);
  const svg = encode(wrapMathmlInSvg(mathml, box.width, box.height));
  const src = mathSource(tex);

  let raster: ResolvedRasterImage | null;
  try {
    raster = await rasterizer.rasterize({
      data: svg,
      format: "svg",
      src,
      intrinsicWidth: box.width,
      intrinsicHeight: box.height,
      targetWidth: box.width,
      targetHeight: box.height,
    });
  } catch (error: unknown) {
    sink.report(
      "image-failed",
      tex,
      display,
      `the rasterizer threw on ${JSON.stringify(abbreviateTex(tex))} ` +
        `(${error instanceof Error ? error.message : String(error)})`,
    );
    return null;
  }

  if (raster === null) {
    sink.report(
      "image-failed",
      tex,
      display,
      `the rasterizer declined ${JSON.stringify(abbreviateTex(tex))}; it may not draw ` +
        `<foreignObject>, which is how MathML is carried inside SVG`,
    );
    return null;
  }

  // The rasteriser measured the equation to draw it, so its aspect ratio is
  // better than our width guess. Keep the estimated height, take the shape.
  const aspect =
    Number.isFinite(raster.width) && Number.isFinite(raster.height) && raster.height > 0
      ? raster.width / raster.height
      : box.width / box.height;
  const height = box.height;
  const width = Math.min(MAX_WIDTH, Math.max(1, Math.round(height * aspect)));

  const resolved: ResolvedImage = {
    format: "svg",
    data: svg,
    width,
    height,
    fallback: raster,
  };

  if (!isValidResolvedImage(resolved)) {
    sink.report(
      "image-failed",
      tex,
      display,
      `the rasterizer returned a raster OOXML cannot hold for ` +
        `${JSON.stringify(abbreviateTex(tex))} (png, jpg, gif or bmp with positive dimensions)`,
    );
    return null;
  }
  return resolved;
}

/** The pseudo-source a math picture carries. Never fetched: it arrives resolved. */
export function mathSource(tex: string): string {
  return `math:${abbreviateTex(tex, 40)}`;
}

/** Builds the image node that replaces an equation in `math: "image"` mode. */
export function mathImageNode(tex: string, resolved: ResolvedImage): ImageNode {
  // alt is the TeX source: it is what a screen reader should say, and what the
  // renderer's placeholder shows if the picture is ever lost.
  return image(mathSource(tex), { alt: tex, resolved });
}
