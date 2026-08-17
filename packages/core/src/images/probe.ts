/**
 * Intrinsic image dimensions, read from the file's own header bytes.
 *
 * ### Why not just use the DOM
 *
 * The obvious way to measure an image is `new Image()` / `createImageBitmap()`.
 * We do not, for four reasons: it needs a DOM (this package must run in a
 * worker, in Node, and on a server), it is asynchronous and racy, it *decodes*
 * the picture (a 10 MiB JPEG becomes ~100 MiB of RGBA in a tab that is already
 * holding a document), and it silently accepts formats OOXML cannot store. A
 * header parse is a few hundred bytes of arithmetic, is synchronous, allocates
 * nothing, and tells us the format as a side effect.
 *
 * ### Why dimensions are mandatory rather than nice-to-have
 *
 * OOXML has no "natural size". Every picture is stamped into the document with
 * an explicit `<wp:extent cx cy>` in EMU, so a renderer that does not know the
 * pixel size has to invent one, and every invented size is wrong by a different
 * amount. Reading them here is what makes {@link import("./layout.js").fitToWidth}
 * aspect-correct.
 *
 * ### EXIF orientation: parsed, reported, deliberately not applied
 *
 * A phone JPEG is stored in sensor order with an EXIF tag saying "rotate me".
 * We read that tag (`orientation`) and expose the rotated size
 * (`orientedWidth`/`orientedHeight`), but the size handed to `ImageRun` is the
 * **stored** one. That is not laziness, it is the only self-consistent choice:
 * the bytes we embed are the unrotated bytes, and Word draws a `a:blip` by
 * scaling the raster into the extent box without consulting EXIF. Swapping the
 * extent would leave the picture just as sideways *and* squashed to the wrong
 * aspect ratio - strictly worse. Callers that genuinely rotate the pixels (a
 * rasteriser, `sharp`, a canvas) should report the post-rotation size
 * themselves; for everyone else we emit an `exif-orientation-ignored` notice so
 * the fact is visible rather than mysterious.
 */

import { hasAscii, hasSignature, readAscii, viewOf } from "./binary.js";
import { probeSvg } from "./svg.js";
import type { ImageProbe, ProbedImageFormat } from "./types.js";

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const;
const GIF87A = [0x47, 0x49, 0x46, 0x38, 0x37, 0x61] as const;
const GIF89A = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61] as const;
const JPEG_SOI = [0xff, 0xd8] as const;
const BMP_SIGNATURE = [0x42, 0x4d] as const;

/** Builds a probe result, filling in the EXIF-rotated size. */
function probeOf(
  format: ProbedImageFormat,
  width: number,
  height: number,
  orientation: number | null = null,
  intrinsic = true,
): ImageProbe | null {
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
  if (width <= 0 || height <= 0) return null;
  const rotated = orientation !== null && orientation >= 5 && orientation <= 8;
  return {
    format,
    width,
    height,
    orientation,
    orientedWidth: rotated ? height : width,
    orientedHeight: rotated ? width : height,
    intrinsic,
  };
}

/* -------------------------------------------------------------------------- */
/* PNG                                                                         */
/* -------------------------------------------------------------------------- */

/**
 * PNG: the 8-byte signature is followed by chunks, and the spec requires IHDR
 * to be the first one. Width and height are big-endian `u32` at its start.
 */
function probePng(bytes: Uint8Array): ImageProbe | null {
  // 8 signature + 4 length + 4 type + 8 dimensions.
  if (bytes.length < 24) return null;
  if (!hasAscii(bytes, 12, "IHDR")) return null;
  const view = viewOf(bytes);
  return probeOf("png", view.getUint32(16), view.getUint32(20));
}

/* -------------------------------------------------------------------------- */
/* GIF                                                                         */
/* -------------------------------------------------------------------------- */

/**
 * GIF: `GIF87a`/`GIF89a` then the Logical Screen Descriptor, whose first four
 * bytes are little-endian width and height.
 */
function probeGif(bytes: Uint8Array): ImageProbe | null {
  if (bytes.length < 10) return null;
  const view = viewOf(bytes);
  return probeOf("gif", view.getUint16(6, true), view.getUint16(8, true));
}

/* -------------------------------------------------------------------------- */
/* BMP                                                                         */
/* -------------------------------------------------------------------------- */

/**
 * BMP: a 14-byte file header, then a DIB header whose *size* field says which
 * of the half-dozen layouts follows. Only the ancient 12-byte
 * `BITMAPCOREHEADER` uses 16-bit dimensions; everything since
 * `BITMAPINFOHEADER` (40) uses signed 32-bit ones, where a negative height
 * means the rows are stored top-down. The sign is a storage detail, so we take
 * the magnitude.
 */
function probeBmp(bytes: Uint8Array): ImageProbe | null {
  if (bytes.length < 18) return null;
  const view = viewOf(bytes);
  const headerSize = view.getUint32(14, true);
  if (headerSize === 12) {
    if (bytes.length < 22) return null;
    return probeOf("bmp", view.getInt16(18, true), Math.abs(view.getInt16(20, true)));
  }
  if (headerSize >= 40) {
    if (bytes.length < 26) return null;
    return probeOf("bmp", view.getInt32(18, true), Math.abs(view.getInt32(22, true)));
  }
  return null;
}

/* -------------------------------------------------------------------------- */
/* WebP                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * WebP: a RIFF container whose first chunk identifies the flavour.
 *
 * - `VP8 ` (lossy) - a 3-byte frame tag, the `9D 01 2A` keyframe start code,
 *   then 16-bit little-endian width and height whose top two bits are an
 *   upscaling hint, not size.
 * - `VP8L` (lossless) - a `0x2F` signature byte then a 32-bit little-endian
 *   bitfield: 14 bits of `width - 1`, 14 bits of `height - 1`.
 * - `VP8X` (extended: animation, alpha, ICC) - a flags byte, 3 reserved bytes,
 *   then 24-bit little-endian `canvasWidth - 1` and `canvasHeight - 1`.
 *
 * Note this is measurement only. `ImageRun` accepts `png | jpg | gif | bmp`, so
 * a WebP still needs an {@link import("./types.js").ImageRasterizer} before it
 * can enter a document; knowing its size is what lets the failure say
 * "unsupported format", not "corrupt file".
 */
function probeWebp(bytes: Uint8Array): ImageProbe | null {
  if (bytes.length < 16) return null;
  if (!hasAscii(bytes, 8, "WEBP")) return null;
  const view = viewOf(bytes);
  const chunk = readAscii(bytes, 12, 4);

  if (chunk === "VP8X") {
    if (bytes.length < 30) return null;
    const width = (view.getUint16(24, true) | (view.getUint8(26) << 16)) + 1;
    const height = (view.getUint16(27, true) | (view.getUint8(29) << 16)) + 1;
    return probeOf("webp", width, height);
  }

  if (chunk === "VP8L") {
    if (bytes.length < 25) return null;
    if (view.getUint8(20) !== 0x2f) return null;
    const bits = view.getUint32(21, true);
    return probeOf("webp", (bits & 0x3fff) + 1, ((bits >> 14) & 0x3fff) + 1);
  }

  if (chunk === "VP8 ") {
    if (bytes.length < 30) return null;
    if (!hasSignature(bytes, [0x9d, 0x01, 0x2a], 23)) return null;
    return probeOf("webp", view.getUint16(26, true) & 0x3fff, view.getUint16(28, true) & 0x3fff);
  }

  return null;
}

/* -------------------------------------------------------------------------- */
/* JPEG                                                                        */
/* -------------------------------------------------------------------------- */

/** Start-of-frame markers. `C4` is Huffman tables, `C8` is reserved, `CC` is arithmetic coding. */
function isStartOfFrame(marker: number): boolean {
  return marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
}

/**
 * Reads the Orientation tag (0x0112) out of an APP1 `Exif\0\0` segment.
 *
 * Just enough TIFF to find one SHORT in IFD0: byte-order mark, magic 42, IFD
 * offset, then a linear scan of 12-byte entries. Sub-IFDs and IFD1 (the
 * thumbnail) are not followed - orientation never lives there.
 */
function readExifOrientation(bytes: Uint8Array, start: number, end: number): number | null {
  if (end - start < 14) return null;
  if (!hasAscii(bytes, start, "Exif")) return null;
  if (bytes[start + 4] !== 0x00 || bytes[start + 5] !== 0x00) return null;

  const tiff = start + 6;
  const view = viewOf(bytes);
  const byteOrder = view.getUint16(tiff, false);
  let little: boolean;
  if (byteOrder === 0x4949) little = true;
  else if (byteOrder === 0x4d4d) little = false;
  else return null;

  if (view.getUint16(tiff + 2, little) !== 42) return null;
  const ifd = tiff + view.getUint32(tiff + 4, little);
  if (ifd < tiff || ifd + 2 > end) return null;

  const entries = view.getUint16(ifd, little);
  for (let i = 0; i < entries; i += 1) {
    const entry = ifd + 2 + i * 12;
    if (entry + 12 > end) return null;
    if (view.getUint16(entry, little) !== 0x0112) continue;
    // Type 3 is SHORT; the value is inlined in the 4-byte value field.
    if (view.getUint16(entry + 2, little) !== 3) return null;
    const value = view.getUint16(entry + 8, little);
    return value >= 1 && value <= 8 ? value : null;
  }
  return null;
}

/**
 * JPEG: there is no header, only a marker stream, so this walks it the way a
 * decoder does.
 *
 * Markers are `FF xx`; `FF` may be repeated as fill. `D0`-`D9` and `01` stand
 * alone, everything else carries a big-endian 16-bit length *including the two
 * length bytes*. Dimensions live in whichever SOF marker the encoder used -
 * baseline (`C0`), progressive (`C2`), lossless, hierarchical - which is why
 * this scans rather than looking at a fixed offset. EXIF (`E1`) is picked up on
 * the way past; it always precedes the SOF.
 */
function probeJpeg(bytes: Uint8Array): ImageProbe | null {
  const view = viewOf(bytes);
  let orientation: number | null = null;
  let offset = 2;

  while (offset + 1 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      // Corrupt or padded stream: resynchronise on the next marker prefix.
      offset += 1;
      continue;
    }
    let marker = bytes[offset + 1] as number;
    while (marker === 0xff && offset + 2 < bytes.length) {
      offset += 1;
      marker = bytes[offset + 1] as number;
    }
    offset += 2;

    // Standalone markers: SOI, TEM and the restart markers carry no payload.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    // EOI, or entropy-coded data we are not going to walk into.
    if (marker === 0xd9 || marker === 0xda) break;

    if (offset + 2 > bytes.length) break;
    const length = view.getUint16(offset, false);
    if (length < 2) return null;
    const segmentStart = offset + 2;
    const segmentEnd = offset + length;
    if (segmentEnd > bytes.length) break; // truncated file

    if (isStartOfFrame(marker)) {
      // precision(1) height(2) width(2)
      if (segmentEnd - segmentStart < 5) return null;
      return probeOf(
        "jpg",
        view.getUint16(segmentStart + 3, false),
        view.getUint16(segmentStart + 1, false),
        orientation,
      );
    }

    if (marker === 0xe1 && orientation === null) {
      orientation = readExifOrientation(bytes, segmentStart, segmentEnd);
    }

    offset = segmentEnd;
  }

  return null;
}

/* -------------------------------------------------------------------------- */
/* Entry point                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Identifies and measures an image from its bytes.
 *
 * Recognises PNG, JPEG, GIF, BMP, WebP and SVG. Returns `null` for anything
 * else, for a header too truncated to trust, and for nonsensical dimensions -
 * and, by contract, **never throws**, including on adversarial input: every
 * read is bounds-checked and the whole body is belt-and-braces wrapped, because
 * this runs on bytes fetched from the internet.
 *
 * @param bytes - The complete file, or at least its first few hundred bytes.
 * @returns What the file is and how big it is, or `null`.
 */
export function probeImage(bytes: Uint8Array): ImageProbe | null {
  try {
    if (bytes.length === 0) return null;
    if (hasSignature(bytes, PNG_SIGNATURE)) return probePng(bytes);
    if (hasSignature(bytes, JPEG_SOI)) return probeJpeg(bytes);
    if (hasSignature(bytes, GIF87A) || hasSignature(bytes, GIF89A)) return probeGif(bytes);
    if (hasSignature(bytes, BMP_SIGNATURE)) return probeBmp(bytes);
    if (hasAscii(bytes, 0, "RIFF")) return probeWebp(bytes);
    return probeSvg(bytes);
  } catch {
    // Unreachable by construction; kept so a future edit cannot turn a
    // malformed download into a rejected conversion.
    return null;
  }
}
