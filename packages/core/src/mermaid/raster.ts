/**
 * SVG (and WebP) to PNG, with the browser's own rasteriser.
 *
 * ### Why this exists at all
 *
 * OOXML cannot hold an SVG on its own. `docx`'s `ImageRun` accepts
 * `type: "svg"`, but the option requires a raster `fallback` alongside it,
 * because the format stores a vector picture as `<a:blip>` (a raster) *plus* an
 * `asvg:svgBlip` extension: Word 2016+ draws the vector, and LibreOffice,
 * Google Docs, Pages, older Word and every thumbnailer draw the raster. There
 * is no legal "vector only" — see `src/images/svg.ts` for the full trap,
 * including the second half of it (a `string` handed to `ImageRun` is treated
 * as base64, so raw markup throws out of the packer).
 *
 * So a diagram has to be rasterised, and the core cannot do it: rasterising
 * needs a canvas, a headless browser or a native library. This module is the
 * browser answer, implementing the {@link ImageRasterizer} interface the image
 * pipeline already defines — which means it is equally usable for ordinary
 * `![](diagram.svg)` images: `createImageResolver({ rasterizer:
 * createCanvasRasterizer() })`.
 *
 * ### The 2x rule, and why WebP is exempt
 *
 * A diagram drawn at exactly its display size looks soft on any screen made
 * this decade and worse in print, so the default renders at **twice** the
 * pixels the picture will occupy. That is free for an SVG: OOXML records the
 * *display* size separately (`<wp:extent>`, in EMU), so the twin can carry as
 * many pixels as it likes — `decodeImage` sizes an SVG from the SVG's own
 * intrinsic dimensions and never from the raster's.
 *
 * WebP is different, and the difference is easy to get wrong: there the raster
 * *replaces* the original, so `decodeImage` uses the raster's pixel dimensions
 * as the display size. Scaling one up by two would silently double the picture
 * in the document. `scale` is therefore applied to SVG only, and WebP is always
 * converted 1:1.
 */

import { parseDataUri } from "../images/data-uri.js";
import { DEFAULT_MAX_BYTES, type ImageRasterizer, type RasterizeRequest } from "../images/types.js";
import type { ResolvedRasterImage } from "../model.js";

/** Options for {@link createCanvasRasterizer}. */
export interface CanvasRasterizerOptions {
  /**
   * Raster pixels per display pixel, for SVG sources. Defaults to `2`.
   *
   * `1` produces a picture that matches its display size exactly; `2` is the
   * usual "retina" choice and what a diagram wants in print. Values are clamped
   * to `[1, 8]`, and further reduced when the result would exceed
   * {@link CanvasRasterizerOptions.maxPixels}.
   *
   * Ignored for WebP; see the note at the top of this file.
   */
  readonly scale?: number | undefined;
  /**
   * Colour to paint behind the picture, as any CSS colour, or `null` to keep
   * the alpha channel. Defaults to `null`.
   *
   * A mermaid diagram is transparent, which is normally right — the Word page
   * shows through. Set `"#ffffff"` when a document might be viewed on a dark
   * page and the diagram's own strokes are dark.
   */
  readonly background?: string | null | undefined;
  /**
   * Hard cap on `width * height` of the raster. Defaults to 16 million (about
   * a 4000x4000 picture, ~64 MB of RGBA in flight).
   *
   * The `scale` is reduced until the raster fits rather than failing: half the
   * resolution is always better than no diagram.
   */
  readonly maxPixels?: number | undefined;
}

/** Highest `scale` that is a sharpening rather than a memory experiment. */
const MAX_SCALE = 8;
const DEFAULT_SCALE = 2;
const DEFAULT_MAX_PIXELS = 16_000_000;

/** The media type to hand a `Blob` for each source format. */
const SOURCE_MEDIA_TYPE: Readonly<Record<"svg" | "webp", string>> = {
  svg: "image/svg+xml;charset=utf-8",
  webp: "image/webp",
};

function finite(value: number | undefined, fallback: number, min: number, max: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, value));
}

/** A positive, finite pixel count, rounded to a whole pixel. */
function pixels(value: number | null | undefined, fallback: number): number {
  if (value === null || value === undefined || !Number.isFinite(value) || value <= 0) {
    return fallback;
  }
  return Math.max(1, Math.round(value));
}

/**
 * The size the picture will occupy, in CSS pixels.
 *
 * `targetWidth`/`targetHeight` are what the image pipeline computed from the
 * page's text column; the intrinsic size is the fallback, and CSS's 300x150
 * replaced-element default the fallback's fallback.
 */
function displaySize(request: RasterizeRequest): { width: number; height: number } {
  const width = pixels(request.targetWidth, pixels(request.intrinsicWidth, 300));
  const height = pixels(request.targetHeight, pixels(request.intrinsicHeight, 150));
  return { width, height };
}

interface CanvasLike {
  width: number;
  height: number;
  getContext(id: "2d"): CanvasContextLike | null;
  toBlob?: (callback: (blob: Blob | null) => void, type?: string) => void;
  toDataURL?: (type?: string) => string;
}

interface CanvasContextLike {
  fillStyle: string;
  fillRect(x: number, y: number, width: number, height: number): void;
  drawImage(image: object, x: number, y: number, width: number, height: number): void;
}

interface ImageLike {
  width: number;
  height: number;
  onload: (() => void) | null;
  onerror: ((error: unknown) => void) | null;
  decoding?: string;
  src: string;
}

/**
 * A `blob:` URL for the source bytes, with a `data:` URL as the fallback.
 *
 * `createObjectURL` is the memory-efficient path and the one every browser has;
 * the fallback exists for the runtimes that have a DOM but no object URLs
 * (some worker/jsdom setups), where percent-encoding the markup still works.
 */
function sourceUrl(request: RasterizeRequest): { url: string; revoke: () => void } {
  const type = SOURCE_MEDIA_TYPE[request.format];
  const canObjectUrl =
    typeof Blob === "function" &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function";

  if (canObjectUrl) {
    // A copy, not a view: the caller's buffer may be a subarray of a larger
    // one, and Blob would otherwise capture the whole thing.
    const blob = new Blob([request.data.slice()], { type });
    const url = URL.createObjectURL(blob);
    return {
      url,
      revoke: () => {
        try {
          URL.revokeObjectURL(url);
        } catch {
          /* nothing to release */
        }
      },
    };
  }

  const text = new TextDecoder().decode(request.data);
  return { url: `data:${type},${encodeURIComponent(text)}`, revoke: () => {} };
}

/** Loads a URL into an `<img>`, resolving when the pixels are available. */
async function loadImage(
  element: ImageLike,
  url: string,
  size: { width: number; height: number },
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    element.onload = () => {
      resolve();
    };
    element.onerror = () => {
      reject(new Error("the browser could not decode the image"));
    };
    // An SVG with no intrinsic size renders at CSS's 300x150 default unless the
    // element is given one; setting both attributes is what makes a viewBox-only
    // diagram rasterise at the size we asked for.
    element.width = size.width;
    element.height = size.height;
    element.decoding = "sync";
    element.src = url;
  });
}

/** Reads the canvas back as PNG bytes, whichever export API it has. */
async function exportPng(canvas: CanvasLike): Promise<Uint8Array> {
  if (typeof canvas.toBlob === "function") {
    const blob = await new Promise<Blob | null>((resolve) => {
      // Typed above; the narrowing is lost through the optional-call syntax.
      (canvas.toBlob as (cb: (value: Blob | null) => void, type?: string) => void)(
        resolve,
        "image/png",
      );
    });
    if (blob === null) throw new Error("canvas.toBlob() produced nothing");
    return new Uint8Array(await blob.arrayBuffer());
  }

  if (typeof canvas.toDataURL === "function") {
    const parsed = parseDataUri(canvas.toDataURL("image/png"), DEFAULT_MAX_BYTES);
    if (!parsed.ok) throw new Error(`canvas.toDataURL() produced ${parsed.message}`);
    return parsed.bytes;
  }

  throw new Error("this canvas has neither toBlob() nor toDataURL()");
}

/**
 * A browser {@link ImageRasterizer}: `<img>` + `<canvas>` + `toBlob`.
 *
 * Every failure is a rejected promise, which the callers of an
 * `ImageRasterizer` (`decodeImage`, and the mermaid pass) already treat as "no
 * raster for this picture" — one placeholder or one kept code fence, never a
 * failed document.
 *
 * @param options - Scale, background and the memory cap. See {@link CanvasRasterizerOptions}.
 * @returns A rasteriser usable both by this module and by
 *   `createImageResolver({ rasterizer })`.
 */
export function createCanvasRasterizer(options: CanvasRasterizerOptions = {}): ImageRasterizer {
  const requestedScale = finite(options.scale, DEFAULT_SCALE, 1, MAX_SCALE);
  const maxPixels = finite(options.maxPixels, DEFAULT_MAX_PIXELS, 1, Number.MAX_SAFE_INTEGER);
  const background = options.background ?? null;

  return {
    async rasterize(request: RasterizeRequest): Promise<ResolvedRasterImage | null> {
      if (typeof document === "undefined") {
        throw new Error("createCanvasRasterizer() needs a DOM; there is no `document` here");
      }

      const display = displaySize(request);
      // WebP replaces the original, so its raster size *is* its display size.
      const wanted = request.format === "webp" ? 1 : requestedScale;
      const fit = Math.sqrt(maxPixels / (display.width * display.height));
      const scale = Math.max(1, Math.min(wanted, Number.isFinite(fit) ? fit : wanted));

      const width = Math.max(1, Math.round(display.width * scale));
      const height = Math.max(1, Math.round(display.height * scale));

      const element = document.createElement("img") as unknown as ImageLike;
      const canvas = document.createElement("canvas") as unknown as CanvasLike;
      const { url, revoke } = sourceUrl(request);

      try {
        await loadImage(element, url, display);

        canvas.width = width;
        canvas.height = height;
        const context = canvas.getContext("2d");
        if (context === null) throw new Error("canvas.getContext('2d') returned null");

        if (background !== null) {
          context.fillStyle = background;
          context.fillRect(0, 0, width, height);
        }
        // Explicit destination size: the browser re-rasterises a vector at the
        // size it is drawn, which is where the extra pixels come from.
        context.drawImage(element as unknown as object, 0, 0, width, height);

        const data = await exportPng(canvas);
        if (data.length === 0) throw new Error("the canvas exported zero bytes");

        return { format: "png", data, width, height };
      } finally {
        revoke();
      }
    },
  };
}
