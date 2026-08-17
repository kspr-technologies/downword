/**
 * Bytes to a {@link ResolvedImage}: the last gate before OOXML.
 *
 * Everything the renderer needs is decided here - the format tag `ImageRun`
 * demands, the pixel dimensions the `<wp:extent>` is computed from, and the
 * raster twin an SVG cannot legally travel without. Anything that would make
 * `ImageRun` produce a picture Word cannot draw is turned into `null` plus a
 * diagnostic, which the renderer already knows how to display as a placeholder.
 *
 * The type sniff is done on the bytes, never on the URL extension or the
 * `Content-Type`: a `.png` that is really a JPEG (extremely common in scraped
 * content) would otherwise be embedded with the wrong `type` and silently fail
 * to render in Word, while an image served as `application/octet-stream` would
 * be rejected for no reason.
 */

import type { RasterImageFormat, ResolvedImage, ResolvedRasterImage } from "../model.js";
import { fitToWidth } from "./layout.js";
import { probeImage } from "./probe.js";
import type { DiagnosticSink, ImageProbe, ImageRasterizer } from "./types.js";

/** The formats `docx`'s `ImageRun` accepts directly. */
const EMBEDDABLE: ReadonlySet<string> = new Set<RasterImageFormat>(["png", "jpg", "gif", "bmp"]);

/** Inputs to {@link decodeImage}. */
export interface DecodeImageOptions {
  /** The source as authored, for diagnostics. */
  readonly src: string;
  /** Where to send diagnostics. */
  readonly sink: DiagnosticSink;
  /** Optional escape hatch for SVG and WebP. See {@link ImageRasterizer}. */
  readonly rasterizer?: ImageRasterizer | null | undefined;
  /**
   * Usable text width in twips, used only to suggest a raster target size.
   * `null` leaves {@link import("./types.js").RasterizeRequest.targetWidth}
   * unset.
   */
  readonly contentWidthTwips?: number | null | undefined;
}

/** Whether a rasteriser's output is something we can actually embed. */
function isUsableRaster(raster: ResolvedRasterImage | null): raster is ResolvedRasterImage {
  if (raster === null) return false;
  if (!EMBEDDABLE.has(raster.format)) return false;
  if (!(raster.data instanceof Uint8Array) || raster.data.length === 0) return false;
  const { width, height } = raster;
  if (!Number.isFinite(width) || !Number.isFinite(height)) return false;
  return width > 0 && height > 0;
}

/** Asks the rasteriser for a raster twin, treating any misbehaviour as a failure. */
async function rasterize(
  bytes: Uint8Array,
  probe: ImageProbe,
  format: "svg" | "webp",
  options: DecodeImageOptions,
): Promise<ResolvedRasterImage | null> {
  const rasterizer = options.rasterizer;
  if (rasterizer === null || rasterizer === undefined) return null;

  const contentWidth = options.contentWidthTwips;
  const target =
    contentWidth !== null && contentWidth !== undefined
      ? fitToWidth({ width: probe.width, height: probe.height }, contentWidth)
      : null;

  let raster: ResolvedRasterImage | null;
  try {
    raster = await rasterizer.rasterize({
      data: bytes,
      format,
      src: options.src,
      intrinsicWidth: probe.intrinsic ? probe.width : null,
      intrinsicHeight: probe.intrinsic ? probe.height : null,
      targetWidth: target?.width ?? null,
      targetHeight: target?.height ?? null,
    });
  } catch {
    // A rasteriser is third-party code (a canvas, a WASM decoder, a subprocess).
    // Its exceptions are this image's problem, not the document's.
    return null;
  }

  return isUsableRaster(raster) ? raster : null;
}

/**
 * Turns bytes into something `ImageRun` can embed.
 *
 * @param bytes - The complete image file.
 * @param options - Source, diagnostics sink and the optional rasteriser.
 * @returns A {@link ResolvedImage}, or `null` when the image cannot be embedded
 *   safely. Never throws.
 */
export async function decodeImage(
  bytes: Uint8Array,
  options: DecodeImageOptions,
): Promise<ResolvedImage | null> {
  const { src, sink } = options;

  if (bytes.length === 0) {
    sink.report("empty", src, `image "${src}" is zero bytes`);
    return null;
  }

  const probe = probeImage(bytes);
  if (probe === null) {
    sink.report(
      "undecodable",
      src,
      `image "${src}" is not a PNG, JPEG, GIF, BMP, WebP or SVG, or its header is truncated`,
    );
    return null;
  }

  if (probe.format === "svg") {
    // The SVG trap: no raster twin, no picture. See svg.ts.
    const fallback = await rasterize(bytes, probe, "svg", options);
    if (fallback === null) {
      const code = options.rasterizer ? "rasterizer-failed" : "svg-no-rasterizer";
      sink.report(
        code,
        src,
        `SVG "${src}" needs a raster fallback to be embedded (OOXML stores both, and readers ` +
          `that cannot draw SVG show the raster); ${
            options.rasterizer
              ? "the supplied rasterizer did not produce one"
              : "no rasterizer was supplied"
          }, so a placeholder was emitted instead`,
      );
      return null;
    }
    return {
      format: "svg",
      // Bytes, never markup: ImageRun runs any string through atob().
      data: bytes,
      width: probe.width,
      height: probe.height,
      fallback,
    };
  }

  if (probe.format === "webp") {
    // Word can display WebP, but the OOXML picture part cannot declare it:
    // ImageRun's type union is png | jpg | gif | bmp. Transcode or bail.
    const raster = await rasterize(bytes, probe, "webp", options);
    if (raster === null) {
      const code = options.rasterizer ? "rasterizer-failed" : "unsupported-format";
      sink.report(
        code,
        src,
        `image "${src}" is a ${probe.width}x${probe.height} WebP, which OOXML cannot embed ` +
          `(png, jpg, gif and bmp only); ${
            options.rasterizer
              ? "the supplied rasterizer failed to convert it"
              : "supply a rasterizer to convert it"
          }`,
      );
      return null;
    }
    return { ...raster, fallback: null };
  }

  if (probe.orientation !== null && probe.orientation >= 5) {
    sink.report(
      "exif-orientation-ignored",
      src,
      `image "${src}" carries EXIF orientation ${probe.orientation}, which rotates it a quarter ` +
        `turn. Word draws the stored pixels and ignores EXIF, so it is embedded at its stored ` +
        `${probe.width}x${probe.height} rather than ${probe.orientedWidth}x${probe.orientedHeight}; ` +
        `rotate the bytes upstream if that matters`,
    );
  }

  return {
    format: probe.format,
    data: bytes,
    width: probe.width,
    height: probe.height,
    fallback: null,
  };
}

/**
 * Whether a {@link ResolvedImage} honours the model's contract.
 *
 * Applied to whatever a third-party
 * {@link import("../render/types.js").ImageResolver} hands back, because the
 * two invariants below are exactly the ones that produce an unopenable `.docx`
 * rather than a visible error: `fallback` is non-`null` **iff** the format is
 * `svg`, and the dimensions are positive finite numbers.
 */
export function isValidResolvedImage(image: ResolvedImage | null): image is ResolvedImage {
  if (image === null) return false;
  if (!(image.data instanceof Uint8Array) || image.data.length === 0) return false;
  const { width, height } = image;
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0)
    return false;
  if (image.format === "svg") return isUsableRaster(image.fallback);
  if (!EMBEDDABLE.has(image.format)) return false;
  return image.fallback === null;
}
