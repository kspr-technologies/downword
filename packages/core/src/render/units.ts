/**
 * OOXML measurement helpers.
 *
 * WordprocessingML mixes four units and gets them wrong silently, so every
 * conversion in the renderer goes through this file:
 *
 * | unit        | where it is used                                  |
 * |-------------|---------------------------------------------------|
 * | twip (1/20 pt) | indents, spacing, page size, table widths      |
 * | half-point  | `w:sz` (font size)                                |
 * | eighth-point| `w:sz` on borders                                 |
 * | CSS pixel   | `ImageRun` transformations (docx multiplies by 9525 to get EMU) |
 */

/** Twentieths of a point. 1 inch = 72 pt = 1440 twips. */
export const TWIPS_PER_INCH = 1440;

/** 1 pt = 20 twips. */
export const TWIPS_PER_POINT = 20;

/**
 * A CSS pixel at 96 dpi is 9525 EMU and a twip is 635 EMU, so one pixel is
 * exactly 15 twips. `ImageRun.transformation` is specified in pixels.
 */
export const TWIPS_PER_PIXEL = 15;

/** CSS reference resolution: 96 pixels to the inch. */
export const PIXELS_PER_INCH = 96;

/** OOXML's base unit: 914400 English Metric Units to the inch. */
export const EMU_PER_INCH = 914400;

/**
 * `914400 / 96` = 9525, exactly.
 *
 * This is the whole reason a picture's size is a *fraction* of a pixel rather
 * than a whole one: `ImageRun.transformation` is in CSS pixels, but docx turns
 * it into `emus: { x: Math.round(px * 9525) }` (9.7.1, `createImageData` in
 * `file/paragraph/run/image-run.ts`) and writes *that* to `<wp:extent>` and
 * `<a:ext>`. The EMU, not the pixel, is the finest size OOXML can express, so
 * rounding to a whole pixel first throws away up to 4762 EMU per axis for
 * nothing. (The sibling `pixels` field is rounded too, and is read by nothing
 * in the library — it is part of `IMediaData` for consumers.)
 */
export const EMU_PER_PIXEL = EMU_PER_INCH / PIXELS_PER_INCH;

/**
 * The smallest picture dimension that survives that conversion.
 *
 * Anything under half an EMU rounds to a `<wp:extent>` of zero, which Word
 * draws as nothing at all.
 */
export const MIN_PICTURE_PIXELS = 1 / EMU_PER_PIXEL;

/** Points -> twips, rounded to the nearest twip. */
export function pointsToTwips(points: number): number {
  return Math.round(points * TWIPS_PER_POINT);
}

/** Points -> half-points (the unit of `w:sz`). */
export function pointsToHalfPoints(points: number): number {
  return Math.round(points * 2);
}

/** Points -> eighth-points (the unit of `w:sz` inside `w:pBdr`/`w:tblBorders`). */
export function pointsToEighthPoints(points: number): number {
  return Math.max(1, Math.round(points * 8));
}

/** Inches -> twips. */
export function inchesToTwips(inches: number): number {
  return Math.round(inches * TWIPS_PER_INCH);
}

/** Twips -> whole CSS pixels, rounded down so an image never overflows. */
export function twipsToPixels(twips: number): number {
  return Math.max(1, Math.floor(twips / TWIPS_PER_PIXEL));
}

/** Clamps `value` into `[min, max]`. */
export function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}
