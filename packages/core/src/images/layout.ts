/**
 * Pixels to OOXML units, and fitting a picture to the text column.
 *
 * ### The conversion, and why it is what it is
 *
 * OOXML sizes a picture in **EMU** (English Metric Units), 914400 to the inch -
 * a number chosen because it divides evenly by both 72 (points) and 96 (CSS
 * pixels), so neither typography nor screen graphics ever round.
 *
 * `docx`'s `ImageRun` does *not* take EMU. Its `transformation: { width,
 * height }` is in **CSS pixels at 96 dpi**, and the library multiplies by 9525
 * on the way out:
 *
 * ```text
 *   914400 EMU/inch ÷ 96 px/inch = 9525 EMU/px
 *   914400 EMU/inch ÷ 1440 twip/inch = 635 EMU/twip   ⇒  1 px = 15 twip
 * ```
 *
 * Every other measurement in the renderer is in twips (page width, margins,
 * indents), so the whole job of this module is that exact-integer 15:1 bridge.
 * Both constants are exact, so a picture's size survives the round trip with no
 * accumulated drift - the only rounding is the deliberate one below.
 *
 * ### Fitting
 *
 * Word does not reflow an oversized picture; it lets it run off the page, past
 * the margin, and out of a table cell. So an image wider than the text column
 * is scaled down, never up (`scale = min(1, …)`), and the height follows the
 * width so the aspect ratio is preserved. The available width is floored into
 * whole pixels first, so a rounding error can only ever make the picture a
 * fifteenth of a twip too *narrow*.
 *
 * {@link fitToWidth} answers that in **whole pixels**, because its job is to
 * size a *raster*: it is the target handed to an
 * {@link import("./types.js").ImageRasterizer}, and what a host that downscales
 * bytes before embedding should aim for. Half a pixel of bitmap does not exist.
 *
 * The renderer applies the same rule (and a second one, for height) at render
 * time in `render/inline.ts:fitToColumn`, and there it deliberately does *not*
 * round: `<wp:extent>` is in EMU, 9525 to the pixel, so quantising to a whole
 * pixel first would distort the aspect ratio by up to half a pixel for nothing.
 * The two are consistent, not conflicting — a picture whose bytes were already
 * rastered to fit needs no further scaling, and `scale` comes out 1.
 */

import {
  EMU_PER_INCH,
  EMU_PER_PIXEL,
  PIXELS_PER_INCH,
  TWIPS_PER_PIXEL,
  twipsToPixels,
} from "../render/units.js";

// All five are defined in `render/units.ts`, where every unit the package
// converts between lives so that each has exactly one definition, and
// re-exported here so a resolver sizing a raster need not import the renderer.
export { EMU_PER_INCH, EMU_PER_PIXEL, PIXELS_PER_INCH, TWIPS_PER_PIXEL, twipsToPixels };

/**
 * A4 with one-inch margins: `11906 - 2 * 1440`.
 *
 * The renderer's default page, restated here so a resolver can compute a raster
 * target without importing the whole render layer.
 */
export const DEFAULT_CONTENT_WIDTH_TWIPS = 9026;

/** A size in CSS pixels. */
export interface PixelSize {
  readonly width: number;
  readonly height: number;
}

/** CSS pixels -> EMU. Exact: 1 px is 9525 EMU. */
export function pixelsToEmu(pixels: number): number {
  return Math.round(pixels * EMU_PER_PIXEL);
}

/** CSS pixels -> twips. Exact: 1 px is 15 twips. */
export function pixelsToTwips(pixels: number): number {
  return Math.round(pixels * TWIPS_PER_PIXEL);
}

/**
 * Scales a picture down to fit `maxWidthTwips`, preserving its aspect ratio.
 *
 * Never enlarges: a 40x40 icon stays 40x40 in a 601-pixel-wide column rather
 * than being blown up to a blurry banner. Both results are whole pixels of at
 * least 1, because a zero-sized `<wp:extent>` makes Word draw nothing at all.
 *
 * @param intrinsic - The picture's own pixel size.
 * @param maxWidthTwips - Usable text width, in twips. Non-positive means "no limit".
 * @returns The size to raster the picture at, in whole CSS pixels.
 */
export function fitToWidth(intrinsic: PixelSize, maxWidthTwips: number): PixelSize {
  const { width, height } = intrinsic;
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return { width: 1, height: 1 };
  }
  if (!Number.isFinite(maxWidthTwips) || maxWidthTwips <= 0) {
    return { width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) };
  }
  const maxWidthPixels = twipsToPixels(maxWidthTwips);
  const scale = Math.min(1, maxWidthPixels / width);
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  };
}
