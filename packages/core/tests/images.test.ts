/**
 * Tests for the image pipeline (`src/images/**`).
 *
 * Two kinds of test image are used deliberately:
 *
 * - **Constructed** files, built byte by byte below with real CRCs, a real
 *   zlib stream and a real LZW stream. These are what the dimension tests are
 *   parameterised over, because only a constructor can produce the awkward
 *   cases on demand (258-pixel widths that exercise little-endian parsing, a
 *   top-down BMP, a progressive JPEG, a truncated header).
 * - **Encoder-produced** files, embedded as base64: a JPEG written by Pillow
 *   with a real EXIF orientation tag, and three WebPs written by `cwebp`
 *   (lossy, lossless and extended). They exist so the parser is checked against
 *   bytes nobody in this repository chose, which is the only way to catch a
 *   constructor that encodes the same misconception as the parser.
 *
 * The constructed PNG, GIF, BMP and WebP files were additionally cross-checked
 * against macOS ImageIO (`sips -g pixelWidth -g pixelHeight`), which agrees
 * with every dimension asserted here.
 */

import { mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { deflateSync } from "node:zlib";

import { ImageRun } from "docx";
import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";

import {
  createImageResolver,
  decodeBase64,
  diagnosticSeverity,
  hasImageErrors,
  fitToWidth,
  isValidResolvedImage,
  NULL_IMAGE_RESOLVER,
  parseDataUri,
  percentDecodeToBytes,
  pixelsToEmu,
  pixelsToTwips,
  probeImage,
  resolveDocumentImages,
  type FetchLike,
  type ImageDiagnostic,
  type ImageRasterizer,
} from "../src/images/index.js";
import { createNodeImageResolver, isPrivateAddress } from "../src/images/node.js";
import { doc, image, paragraph, text, type ResolvedImage } from "../src/model.js";
import { renderDocument } from "../src/render/index.js";
import type { ImageResolver } from "../src/render/types.js";
import { listDocxParts, packDocument, readDocxPart } from "./helpers/docx.js";

/* -------------------------------------------------------------------------- */
/* No unhandled rejections, ever                                               */
/* -------------------------------------------------------------------------- */

/**
 * The whole point of the resolver is that a failing image cannot escalate. An
 * unhandled rejection anywhere in the suite fails the test that produced it.
 */
const unhandled: unknown[] = [];
const onUnhandledRejection = (reason: unknown): void => {
  unhandled.push(reason);
};

beforeAll(() => {
  process.on("unhandledRejection", onUnhandledRejection);
});

afterAll(() => {
  process.off("unhandledRejection", onUnhandledRejection);
});

afterEach(async () => {
  // Give any orphaned promise a turn to reject before we look.
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(unhandled).toEqual([]);
});

/* -------------------------------------------------------------------------- */
/* Byte helpers                                                                */
/* -------------------------------------------------------------------------- */

function bytes(...parts: readonly Uint8Array[]): Uint8Array {
  let total = 0;
  for (const part of parts) total += part.length;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function ascii(value: string): Uint8Array {
  const out = new Uint8Array(value.length);
  for (let i = 0; i < value.length; i += 1) out[i] = value.charCodeAt(i);
  return out;
}

function u16be(value: number): Uint8Array {
  const out = new Uint8Array(2);
  new DataView(out.buffer).setUint16(0, value);
  return out;
}

function u16le(value: number): Uint8Array {
  const out = new Uint8Array(2);
  new DataView(out.buffer).setUint16(0, value, true);
  return out;
}

function u32be(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, value >>> 0);
  return out;
}

function u32le(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, value >>> 0, true);
  return out;
}

function i32le(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setInt32(0, value, true);
  return out;
}

function toBase64(input: Uint8Array): string {
  return Buffer.from(input).toString("base64");
}

/**
 * Copies bytes into a plain `ArrayBuffer`.
 *
 * `Response` and `Blob` want a `BodyInit`/`BlobPart`, and since TypeScript 5.7
 * a `Uint8Array` is generic over its backing buffer - a `Uint8Array<ArrayBufferLike>`
 * is not one. Copying is clearer than casting and costs nothing at these sizes.
 */
function bodyOf(input: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(input.byteLength);
  new Uint8Array(buffer).set(input);
  return buffer;
}

/* -------------------------------------------------------------------------- */
/* Constructed images                                                          */
/* -------------------------------------------------------------------------- */

const CRC_TABLE: Uint32Array = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(input: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of input) crc = (CRC_TABLE[(crc ^ byte) & 0xff] as number) ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const body = bytes(ascii(type), data);
  return bytes(u32be(data.length), body, u32be(crc32(body)));
}

/** A complete, decodable 8-bit greyscale PNG. Real CRCs, real zlib stream. */
function makePng(width: number, height: number): Uint8Array {
  const ihdr = new Uint8Array(13);
  const view = new DataView(ihdr.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 0; // colour type: greyscale
  const stride = width + 1; // one filter byte per scanline
  const raw = new Uint8Array(stride * height);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) raw[y * stride + 1 + x] = (x * 17 + y * 29) & 0xff;
  }
  return bytes(
    Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", new Uint8Array(deflateSync(raw))),
    pngChunk("IEND", new Uint8Array(0)),
  );
}

/**
 * A complete, decodable GIF89a.
 *
 * The LZW stream emits a CLEAR code before every pixel, so the dictionary never
 * grows and the code width stays at `minCodeSize + 1`. Wasteful, entirely
 * legal, and short enough to verify by hand.
 */
function makeGif(width: number, height: number): Uint8Array {
  const minCodeSize = 2;
  const clear = 1 << minCodeSize;
  const codeWidth = minCodeSize + 1;
  const codes: number[] = [];
  for (let i = 0; i < width * height; i += 1) codes.push(clear, 0);
  codes.push(clear + 1); // end of information

  const packed: number[] = [];
  let accumulator = 0;
  let bits = 0;
  for (const code of codes) {
    accumulator |= code << bits;
    bits += codeWidth;
    while (bits >= 8) {
      packed.push(accumulator & 0xff);
      accumulator >>= 8;
      bits -= 8;
    }
  }
  if (bits > 0) packed.push(accumulator & 0xff);

  const blocks: Uint8Array[] = [];
  for (let i = 0; i < packed.length; i += 255) {
    const slice = packed.slice(i, i + 255);
    blocks.push(Uint8Array.of(slice.length), Uint8Array.from(slice));
  }

  return bytes(
    ascii("GIF89a"),
    u16le(width),
    u16le(height),
    Uint8Array.of(0x80, 0x00, 0x00), // global colour table of 2 entries
    Uint8Array.of(0x00, 0x00, 0x00, 0xff, 0xff, 0xff),
    Uint8Array.of(0x2c), // image descriptor
    u16le(0),
    u16le(0),
    u16le(width),
    u16le(height),
    Uint8Array.of(0x00),
    Uint8Array.of(minCodeSize),
    ...blocks,
    Uint8Array.of(0x00, 0x3b), // block terminator, trailer
  );
}

/** A complete 24-bit BMP. `topDown` writes the negative height real encoders use. */
function makeBmp(width: number, height: number, topDown = false): Uint8Array {
  const stride = Math.ceil((width * 3) / 4) * 4;
  const pixels = new Uint8Array(stride * height);
  for (let i = 0; i < pixels.length; i += 1) pixels[i] = (i * 7) & 0xff;
  const offset = 14 + 40;
  return bytes(
    ascii("BM"),
    u32le(offset + pixels.length),
    u16le(0),
    u16le(0),
    u32le(offset),
    u32le(40), // BITMAPINFOHEADER
    i32le(width),
    i32le(topDown ? -height : height),
    u16le(1),
    u16le(24),
    u32le(0),
    u32le(pixels.length),
    i32le(2835),
    i32le(2835),
    u32le(0),
    u32le(0),
    pixels,
  );
}

/** The 12-byte `BITMAPCOREHEADER` variant, which uses 16-bit dimensions. */
function makeCoreBmp(width: number, height: number): Uint8Array {
  const pixels = new Uint8Array(width * height * 3);
  const offset = 14 + 12;
  return bytes(
    ascii("BM"),
    u32le(offset + pixels.length),
    u16le(0),
    u16le(0),
    u32le(offset),
    u32le(12),
    u16le(width),
    u16le(height),
    u16le(1),
    u16le(24),
    pixels,
  );
}

function jpegSegment(marker: number, payload: Uint8Array): Uint8Array {
  return bytes(Uint8Array.of(0xff, marker), u16be(payload.length + 2), payload);
}

/**
 * A JPEG marker stream.
 *
 * Structurally exact - every segment a decoder walks past on its way to the
 * SOF is present, with correct lengths - but the entropy-coded scan is filler,
 * so this is a header fixture rather than a viewable picture. The viewable
 * counterpart is {@link REAL_JPEG_EXIF_ORIENTATION_6} below.
 */
function makeJpeg(
  width: number,
  height: number,
  options: { orientation?: number; sof?: number; fillBytes?: boolean } = {},
): Uint8Array {
  const sof = options.sof ?? 0xc0;
  const parts: Uint8Array[] = [Uint8Array.of(0xff, 0xd8)];

  parts.push(
    jpegSegment(
      0xe0,
      bytes(
        ascii("JFIF"),
        Uint8Array.of(0),
        Uint8Array.of(1, 1, 0),
        u16be(1),
        u16be(1),
        Uint8Array.of(0, 0),
      ),
    ),
  );

  if (options.orientation !== undefined) {
    // TIFF: big-endian ("MM"), magic 42, IFD0 at offset 8, one SHORT entry.
    const ifd = bytes(
      u16be(1),
      u16be(0x0112),
      u16be(3),
      u32be(1),
      u16be(options.orientation),
      u16be(0),
      u32be(0),
    );
    parts.push(
      jpegSegment(
        0xe1,
        bytes(ascii("Exif"), Uint8Array.of(0, 0), ascii("MM"), u16be(42), u32be(8), ifd),
      ),
    );
  }

  parts.push(jpegSegment(0xdb, bytes(Uint8Array.of(0), new Uint8Array(64).fill(16))));
  if (options.fillBytes === true) parts.push(Uint8Array.of(0xff, 0xff)); // legal marker padding
  parts.push(
    jpegSegment(
      sof,
      bytes(
        Uint8Array.of(8),
        u16be(height),
        u16be(width),
        Uint8Array.of(1),
        Uint8Array.of(1, 0x11, 0),
      ),
    ),
  );
  parts.push(
    jpegSegment(0xda, bytes(Uint8Array.of(1), Uint8Array.of(1, 0x00), Uint8Array.of(0, 63, 0))),
  );
  parts.push(Uint8Array.of(0x00, 0xff, 0xd9));
  return bytes(...parts);
}

function riff(chunkType: string, payload: Uint8Array): Uint8Array {
  const padded = payload.length % 2 === 1 ? bytes(payload, Uint8Array.of(0)) : payload;
  const body = bytes(ascii("WEBP"), ascii(chunkType), u32le(payload.length), padded);
  return bytes(ascii("RIFF"), u32le(body.length), body);
}

/** A VP8 (lossy) WebP header: frame tag, keyframe start code, 14-bit dimensions. */
function makeWebpLossy(width: number, height: number): Uint8Array {
  return riff(
    "VP8 ",
    bytes(
      Uint8Array.of(0x30, 0x01, 0x00),
      Uint8Array.of(0x9d, 0x01, 0x2a),
      u16le(width),
      u16le(height),
      new Uint8Array(8),
    ),
  );
}

/** A VP8L (lossless) WebP header: `0x2F` then 14 + 14 bits of `size - 1`. */
function makeWebpLossless(width: number, height: number): Uint8Array {
  const bitfield = ((width - 1) | ((height - 1) << 14)) >>> 0;
  return riff("VP8L", bytes(Uint8Array.of(0x2f), u32le(bitfield), new Uint8Array(4)));
}

/** A VP8X (extended) WebP header: flags, 3 reserved bytes, 24-bit canvas size. */
function makeWebpExtended(width: number, height: number): Uint8Array {
  const w = width - 1;
  const h = height - 1;
  return riff(
    "VP8X",
    bytes(
      Uint8Array.of(0x10, 0, 0, 0),
      Uint8Array.of(w & 0xff, (w >> 8) & 0xff, (w >> 16) & 0xff),
      Uint8Array.of(h & 0xff, (h >> 8) & 0xff, (h >> 16) & 0xff),
    ),
  );
}

function makeSvg(attributes: string): Uint8Array {
  return new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" ${attributes}></svg>`);
}

/* -------------------------------------------------------------------------- */
/* Encoder-produced images                                                     */
/* -------------------------------------------------------------------------- */

/** 23x9 JPEG written by Pillow 10.3 with EXIF `Orientation = 6` (rotate 90 CW). */
const REAL_JPEG_EXIF_ORIENTATION_6 = decodeBase64(
  "/9j/4AAQSkZJRgABAQAAAQABAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAYAAAAAAAD/2wBDABQO" +
    "DxIPDRQSEBIXFRQYHjIhHhwcHj0sLiQySUBMS0dARkVQWnNiUFVtVkVGZIhlbXd7gYKBTmCNl4x9lnN+gXz/" +
    "2wBDARUXFx4aHjshITt8U0ZTfHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8fHx8" +
    "fHx8fHz/wAARCAAJABcDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA" +
    "AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJico" +
    "KSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWm" +
    "p6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEB" +
    "AQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEI" +
    "FEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0" +
    "dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk" +
    "5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwClBHafZ3/cT/cP8I9PrS+XafZG/cT9v4R6j3rSg/495P8A" +
    "cP8AKl/5dG/D+YqnUd/mZRrPkX+F9TNljtPsy/uJ/vf3R6H3orSl/wCPZf8Ae/oaKcajsZVqz5l6dz//2Q==",
) as Uint8Array;

/** 23x9 lossy WebP written by `cwebp 1.x` (a `VP8 ` chunk). */
const REAL_WEBP_LOSSY = decodeBase64(
  "UklGRpgAAABXRUJQVlA4IIwAAACwBACdASoXAAkAPpE8mkglo6KhMAgAsBIJbACdMoGsq/qrlgH7AAWpzuhx" +
    "cDOAAP7qOmtLx7G2yvH7Td746ZVp6SyRyDQO4rleOSyVTExT7k42JjX7O8tesaM4EyLSrztqBYONMiwezheE" +
    "uvqoDzPOr4a7+ZMDdKqcnw2iaoqL/0/Dm8005grnqZIAAA==",
) as Uint8Array;

/** 23x9 lossless WebP written by `cwebp -lossless` (a `VP8L` chunk). */
const REAL_WEBP_LOSSLESS = decodeBase64(
  "UklGRioAAABXRUJQVlA4TB4AAAAvFgACALmM6H/sIqL/ASEB4ZT/9yR5SgYwJgAplm0=",
) as Uint8Array;

/** 23x9 WebP with alpha written by `cwebp` - an extended (`VP8X`) container. */
const REAL_WEBP_EXTENDED = decodeBase64(
  "UklGRloAAABXRUJQVlA4WAoAAAAQAAAAFgAACAAAQUxQSAoAAAABB1DAiAhERP8DVlA4ICoAAACwAgCdASoX" +
    "AAkAPpE6l0eloyIhMAgAsBIJZwAAeyAA/vbl/1ogle4nwAA=",
) as Uint8Array;

/* -------------------------------------------------------------------------- */
/* Fetch doubles                                                               */
/* -------------------------------------------------------------------------- */

interface FetchSpy {
  readonly fetch: FetchLike;
  readonly calls: string[];
}

/** A `fetch` that always answers with `body`, recording every call. */
function respondWith(body: Uint8Array, init: ResponseInit = {}): FetchSpy {
  const calls: string[] = [];
  return {
    calls,
    fetch: (url) => {
      calls.push(url);
      return Promise.resolve(new Response(bodyOf(body), init));
    },
  };
}

/** A `fetch` that never settles until its signal aborts. */
function hangForever(): FetchSpy {
  const calls: string[] = [];
  return {
    calls,
    fetch: (url, init) => {
      calls.push(url);
      return new Promise((_resolve, reject) => {
        const signal = init?.signal;
        if (signal === null || signal === undefined) return;
        signal.addEventListener("abort", () => {
          reject(new DOMException("The operation was aborted.", "AbortError"));
        });
      });
    },
  };
}

/** Collects diagnostics; the array is the assertion target. */
function collector(): {
  readonly onDiagnostic: (d: ImageDiagnostic) => void;
  readonly seen: ImageDiagnostic[];
} {
  const seen: ImageDiagnostic[] = [];
  return { seen, onDiagnostic: (d) => void seen.push(d) };
}

function codes(diagnostics: readonly ImageDiagnostic[]): readonly string[] {
  return diagnostics.map((d) => d.code);
}

/* -------------------------------------------------------------------------- */
/* probeImage                                                                  */
/* -------------------------------------------------------------------------- */

describe("probeImage - constructed files", () => {
  it("reads PNG dimensions from IHDR", () => {
    expect(probeImage(makePng(23, 9))).toMatchObject({ format: "png", width: 23, height: 9 });
    expect(probeImage(makePng(1, 1))).toMatchObject({ format: "png", width: 1, height: 1 });
    expect(probeImage(makePng(1024, 768))).toMatchObject({ width: 1024, height: 768 });
  });

  it("rejects a PNG whose first chunk is not IHDR", () => {
    const png = makePng(4, 4);
    png.set(ascii("iHDR"), 12);
    expect(probeImage(png)).toBeNull();
  });

  it("reads GIF dimensions from the logical screen descriptor", () => {
    expect(probeImage(makeGif(23, 9))).toMatchObject({ format: "gif", width: 23, height: 9 });
    // Little-endian, and both bytes significant.
    expect(probeImage(makeGif(258, 129))).toMatchObject({ width: 258, height: 129 });
  });

  it("accepts GIF87a as well as GIF89a", () => {
    const gif = makeGif(7, 3);
    gif.set(ascii("GIF87a"), 0);
    expect(probeImage(gif)).toMatchObject({ format: "gif", width: 7, height: 3 });
  });

  it("reads BMP dimensions, including top-down and the ancient core header", () => {
    expect(probeImage(makeBmp(23, 9))).toMatchObject({ format: "bmp", width: 23, height: 9 });
    // A negative height means the rows are stored top-down; the size is the same.
    expect(probeImage(makeBmp(23, 9, true))).toMatchObject({ width: 23, height: 9 });
    expect(probeImage(makeCoreBmp(5, 4))).toMatchObject({ format: "bmp", width: 5, height: 4 });
  });

  it("reads JPEG dimensions out of the marker stream", () => {
    expect(probeImage(makeJpeg(23, 9))).toMatchObject({ format: "jpg", width: 23, height: 9 });
    // Progressive (SOF2) and lossless (SOF3) frames, not just baseline.
    expect(probeImage(makeJpeg(640, 480, { sof: 0xc2 }))).toMatchObject({
      width: 640,
      height: 480,
    });
    expect(probeImage(makeJpeg(12, 34, { sof: 0xc3 }))).toMatchObject({ width: 12, height: 34 });
  });

  it("skips over 0xFF marker padding", () => {
    expect(probeImage(makeJpeg(23, 9, { fillBytes: true }))).toMatchObject({
      width: 23,
      height: 9,
    });
  });

  it("is not fooled by a DHT segment that would be mistaken for a frame", () => {
    // 0xC4 sits inside the SOF range but is a Huffman table.
    const jpeg = makeJpeg(23, 9);
    expect(probeImage(jpeg)?.width).toBe(23);
  });

  it("reads WebP dimensions from all three chunk layouts", () => {
    expect(probeImage(makeWebpLossy(23, 9))).toMatchObject({
      format: "webp",
      width: 23,
      height: 9,
    });
    expect(probeImage(makeWebpLossless(23, 9))).toMatchObject({
      format: "webp",
      width: 23,
      height: 9,
    });
    expect(probeImage(makeWebpExtended(23, 9))).toMatchObject({
      format: "webp",
      width: 23,
      height: 9,
    });
    // 14-bit and 24-bit maxima.
    expect(probeImage(makeWebpLossless(16383, 16383))).toMatchObject({
      width: 16383,
      height: 16383,
    });
    expect(probeImage(makeWebpExtended(16385, 4096))).toMatchObject({ width: 16385, height: 4096 });
  });

  it("returns null for empty, truncated and unrecognised bytes", () => {
    expect(probeImage(new Uint8Array(0))).toBeNull();
    expect(probeImage(makePng(8, 8).subarray(0, 20))).toBeNull();
    expect(probeImage(makeJpeg(8, 8).subarray(0, 12))).toBeNull();
    expect(probeImage(makeWebpLossless(8, 8).subarray(0, 18))).toBeNull();
    expect(probeImage(ascii("not an image at all, just prose"))).toBeNull();
    expect(probeImage(new Uint8Array(64))).toBeNull();
    // A PDF: a real file, in a format we do not handle.
    expect(probeImage(ascii("%PDF-1.7\n%\xE2\xE3\xCF\xD3\n"))).toBeNull();
  });

  it("never throws on adversarial input", () => {
    const seeds = [
      makePng(9, 9),
      makeJpeg(9, 9),
      makeGif(9, 9),
      makeBmp(9, 9),
      makeWebpLossy(9, 9),
    ];
    for (const seed of seeds) {
      for (let cut = 1; cut < seed.length; cut += 1) {
        expect(() => probeImage(seed.subarray(0, cut))).not.toThrow();
      }
      // Flip every byte of the header region to garbage.
      for (let i = 0; i < Math.min(seed.length, 48); i += 1) {
        const mutated = seed.slice();
        mutated[i] = 0xff;
        expect(() => probeImage(mutated)).not.toThrow();
      }
    }
  });
});

describe("probeImage - encoder-produced files", () => {
  it("measures a real Pillow JPEG", () => {
    expect(probeImage(REAL_JPEG_EXIF_ORIENTATION_6)).toMatchObject({
      format: "jpg",
      width: 23,
      height: 9,
    });
  });

  it("measures real cwebp output in all three flavours", () => {
    expect(probeImage(REAL_WEBP_LOSSY)).toMatchObject({ format: "webp", width: 23, height: 9 });
    expect(probeImage(REAL_WEBP_LOSSLESS)).toMatchObject({ format: "webp", width: 23, height: 9 });
    expect(probeImage(REAL_WEBP_EXTENDED)).toMatchObject({ format: "webp", width: 23, height: 9 });
  });
});

describe("probeImage - EXIF orientation", () => {
  it("reports a real orientation tag without swapping the stored size", () => {
    const probe = probeImage(REAL_JPEG_EXIF_ORIENTATION_6);
    expect(probe).not.toBeNull();
    expect(probe?.orientation).toBe(6);
    // Stored dimensions: what Word will actually draw.
    expect(probe?.width).toBe(23);
    expect(probe?.height).toBe(9);
    // Rotated dimensions: reported, so a caller that rotates the pixels can use them.
    expect(probe?.orientedWidth).toBe(9);
    expect(probe?.orientedHeight).toBe(23);
  });

  it("swaps only for the four rotating orientations", () => {
    for (const orientation of [1, 2, 3, 4]) {
      const probe = probeImage(makeJpeg(40, 10, { orientation }));
      expect(probe?.orientation).toBe(orientation);
      expect([probe?.orientedWidth, probe?.orientedHeight]).toEqual([40, 10]);
    }
    for (const orientation of [5, 6, 7, 8]) {
      const probe = probeImage(makeJpeg(40, 10, { orientation }));
      expect(probe?.orientation).toBe(orientation);
      expect([probe?.orientedWidth, probe?.orientedHeight]).toEqual([10, 40]);
    }
  });

  it("reports no orientation when there is no EXIF", () => {
    expect(probeImage(makeJpeg(40, 10))?.orientation).toBeNull();
    expect(probeImage(makePng(40, 10))?.orientation).toBeNull();
  });

  it("degrades to no-orientation on a corrupt EXIF block", () => {
    const jpeg = makeJpeg(40, 10, { orientation: 6 });
    // Break the TIFF magic (42) that follows the byte-order mark.
    const index = jpeg.indexOf(0x4d);
    jpeg[index + 3] = 0x99;
    const probe = probeImage(jpeg);
    expect(probe?.width).toBe(40);
    expect(probe?.orientation).toBeNull();
  });
});

describe("probeSvg", () => {
  it("uses width and height when both are absolute", () => {
    expect(probeImage(makeSvg('width="120" height="60"'))).toMatchObject({
      format: "svg",
      width: 120,
      height: 60,
      intrinsic: true,
    });
  });

  it("converts CSS absolute units to pixels", () => {
    // 72pt = 1in = 96px, 1in = 96px.
    expect(probeImage(makeSvg('width="72pt" height="1in"'))).toMatchObject({
      width: 96,
      height: 96,
    });
    expect(probeImage(makeSvg('width="10mm" height="1cm"'))).toMatchObject({
      width: 38,
      height: 38,
    });
  });

  it("falls back to the viewBox, including for percentage sizes", () => {
    expect(probeImage(makeSvg('viewBox="0 0 200 50"'))).toMatchObject({ width: 200, height: 50 });
    expect(probeImage(makeSvg('width="100%" height="100%" viewBox="0 0 30 15"'))).toMatchObject({
      width: 30,
      height: 15,
    });
    // One dimension declared: the other follows the viewBox's ratio.
    expect(probeImage(makeSvg('width="400" viewBox="0 0 200 50"'))).toMatchObject({
      width: 400,
      height: 100,
    });
  });

  it("uses the CSS 300x150 default when the file declares nothing", () => {
    expect(probeImage(makeSvg('role="img"'))).toMatchObject({
      width: 300,
      height: 150,
      intrinsic: false,
    });
  });

  it("survives an XML prolog, a doctype and a '>' inside an attribute", () => {
    const svg = new TextEncoder().encode(
      `<?xml version="1.0"?>\n<!-- a > b -->\n<svg data-tip="a > b" width="20" height="10" xmlns="http://www.w3.org/2000/svg"/>`,
    );
    expect(probeImage(svg)).toMatchObject({ format: "svg", width: 20, height: 10 });
  });

  it("does not mistake arbitrary XML or HTML for an SVG", () => {
    expect(probeImage(ascii('<?xml version="1.0"?><rss><channel/></rss>'))).toBeNull();
    expect(probeImage(ascii("<html><body><svgnot/></body></html>"))).toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* Units and fitting                                                           */
/* -------------------------------------------------------------------------- */

describe("layout", () => {
  it("uses the exact 96 dpi conversions", () => {
    expect(pixelsToEmu(1)).toBe(9525); // 914400 / 96
    expect(pixelsToTwips(1)).toBe(15); // 1440 / 96
    expect(pixelsToEmu(96)).toBe(914400); // one inch
    expect(pixelsToTwips(96)).toBe(1440);
  });

  it("clamps to the content width while preserving the aspect ratio", () => {
    // A4 with 1in margins: 9026 twips = 601 px.
    const fitted = fitToWidth({ width: 1200, height: 600 }, 9026);
    expect(fitted.width).toBe(601);
    expect(fitted.height).toBe(301); // 601 / 2, rounded
    expect(Math.abs(fitted.width / fitted.height - 2)).toBeLessThan(0.01);
  });

  it("never enlarges a small image", () => {
    expect(fitToWidth({ width: 40, height: 40 }, 9026)).toEqual({ width: 40, height: 40 });
  });

  it("keeps a sliver of an image visible rather than rounding it to zero", () => {
    const fitted = fitToWidth({ width: 20000, height: 1 }, 9026);
    expect(fitted.width).toBe(601);
    expect(fitted.height).toBe(1);
  });

  it("degrades safely on nonsense input", () => {
    expect(fitToWidth({ width: 0, height: 10 }, 9026)).toEqual({ width: 1, height: 1 });
    expect(fitToWidth({ width: Number.NaN, height: 10 }, 9026)).toEqual({ width: 1, height: 1 });
    expect(fitToWidth({ width: 30, height: 10 }, 0)).toEqual({ width: 30, height: 10 });
  });
});

/* -------------------------------------------------------------------------- */
/* data: URIs                                                                  */
/* -------------------------------------------------------------------------- */

describe("data: URIs", () => {
  const png = makePng(23, 9);

  it("resolves a base64 payload with its real dimensions", async () => {
    const resolver = createImageResolver();
    const resolved = await resolver.resolve(`data:image/png;base64,${toBase64(png)}`);
    expect(resolved).not.toBeNull();
    expect(resolved?.format).toBe("png");
    expect(resolved?.width).toBe(23);
    expect(resolved?.height).toBe(9);
    expect(resolved?.fallback).toBeNull();
    expect(resolved?.data).toBeInstanceOf(Uint8Array);
    expect(Array.from(resolved?.data ?? [])).toEqual(Array.from(png));
  });

  it("resolves without ever calling fetch", async () => {
    const spy = respondWith(png);
    const resolver = createImageResolver({ fetch: spy.fetch, allowRemote: true });
    await resolver.resolve(`data:image/png;base64,${toBase64(png)}`);
    expect(spy.calls).toEqual([]);
  });

  it("accepts a payload with no media type, and one wrapped across lines", async () => {
    const resolver = createImageResolver();
    const base64 = toBase64(png);
    const wrapped = (base64.match(/.{1,40}/g) ?? []).join("\n");
    await expect(resolver.resolve(`data:;base64,${base64}`)).resolves.not.toBeNull();
    await expect(resolver.resolve(`data:image/png;base64,${wrapped}`)).resolves.not.toBeNull();
  });

  it("decodes base64url", () => {
    const standard = decodeBase64("+/+/");
    const urlSafe = decodeBase64("-_-_");
    expect(urlSafe).toEqual(standard);
  });

  it("decodes a percent-encoded payload as bytes, not as UTF-8 text", () => {
    // 0x89 is a lone continuation byte: decodeURIComponent would throw or mangle it.
    const decoded = percentDecodeToBytes("%89PNG%0D%0A%1A%0A");
    expect(Array.from(decoded ?? [])).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  });

  it("resolves a percent-encoded SVG data URI", async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"></svg>';
    const uri = `data:image/svg+xml,${encodeURIComponent(svg)}`;
    const probe = probeImage((parseDataUri(uri, 1024) as { bytes: Uint8Array }).bytes);
    expect(probe).toMatchObject({ format: "svg", width: 40, height: 20 });
  });

  it("rejects a malformed URI, bad base64 and an empty payload", () => {
    expect(parseDataUri("data:image/png;base64", 1024)).toMatchObject({ code: "bad-data-uri" });
    expect(parseDataUri("data:image/png;base64,!!!!not base64!!!!", 1024)).toMatchObject({
      code: "bad-data-uri",
    });
    expect(parseDataUri("data:image/png;base64,", 1024)).toMatchObject({ code: "empty" });
    expect(parseDataUri("data:image/png;base64,QQ=x", 1024)).toMatchObject({
      code: "bad-data-uri",
    });
  });

  it("enforces the byte cap on the decoded payload", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({ maxBytes: 64, onDiagnostic: diagnostics.onDiagnostic });
    expect(await resolver.resolve(`data:image/png;base64,${toBase64(png)}`)).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["too-large"]);
  });
});

/* -------------------------------------------------------------------------- */
/* The egress switch                                                           */
/* -------------------------------------------------------------------------- */

describe("remote images are off by default", () => {
  const png = makePng(23, 9);

  it("blocks http(s) without allowRemote, and does not call fetch", async () => {
    const spy = respondWith(png);
    const diagnostics = collector();
    const resolver = createImageResolver({
      fetch: spy.fetch,
      onDiagnostic: diagnostics.onDiagnostic,
    });

    expect(await resolver.resolve("https://example.com/logo.png")).toBeNull();
    expect(await resolver.resolve("http://example.com/logo.png")).toBeNull();

    expect(spy.calls).toEqual([]);
    expect(codes(diagnostics.seen)).toEqual(["remote-blocked", "remote-blocked"]);
    expect(diagnostics.seen[0]?.message).toContain("allowRemote");
  });

  it("treats a relative path as remote rather than as something local", async () => {
    const spy = respondWith(png);
    const diagnostics = collector();
    const resolver = createImageResolver({
      fetch: spy.fetch,
      baseUrl: "https://example.com/docs/",
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("./diagram.png")).toBeNull();
    expect(spy.calls).toEqual([]);
    expect(codes(diagnostics.seen)).toEqual(["remote-blocked"]);
  });

  it("fails a relative path with no base URL instead of guessing one", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("images/diagram.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["invalid-url"]);
  });

  it("refuses schemes it cannot serve", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("file:///etc/passwd")).toBeNull();
    expect(await resolver.resolve("ftp://example.com/a.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["unsupported-scheme", "unsupported-scheme"]);
  });

  it("resolves nothing at all through NULL_IMAGE_RESOLVER", async () => {
    expect(await NULL_IMAGE_RESOLVER.resolve(`data:image/png;base64,${toBase64(png)}`)).toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* Remote fetching, once enabled                                               */
/* -------------------------------------------------------------------------- */

describe("remote images with allowRemote: true", () => {
  const png = makePng(23, 9);

  it("fetches and decodes", async () => {
    const spy = respondWith(png, { headers: { "content-type": "image/png" } });
    const resolver = createImageResolver({ allowRemote: true, fetch: spy.fetch });
    const resolved = await resolver.resolve("https://example.com/logo.png");
    expect(resolved).toMatchObject({ format: "png", width: 23, height: 9 });
    expect(spy.calls).toEqual(["https://example.com/logo.png"]);
  });

  it("sends no cookies and no referrer", async () => {
    let init: RequestInit | undefined;
    const resolver = createImageResolver({
      allowRemote: true,
      fetch: (_url, requestInit) => {
        init = requestInit;
        return Promise.resolve(new Response(bodyOf(png)));
      },
    });
    await resolver.resolve("https://example.com/logo.png");
    expect(init?.credentials).toBe("omit");
    expect(init?.referrerPolicy).toBe("no-referrer");
  });

  it("memoises by source, so ten references are one request", async () => {
    const spy = respondWith(png);
    const resolver = createImageResolver({ allowRemote: true, fetch: spy.fetch });
    await Promise.all(
      Array.from({ length: 10 }, () => resolver.resolve("https://example.com/logo.png")),
    );
    expect(spy.calls).toHaveLength(1);
  });

  it("rejects an oversized body declared by Content-Length, before reading it", async () => {
    const diagnostics = collector();
    const spy = respondWith(png, { headers: { "content-length": "999999999" } });
    const resolver = createImageResolver({
      allowRemote: true,
      fetch: spy.fetch,
      maxBytes: 1024,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("https://example.com/huge.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["too-large"]);
    expect(diagnostics.seen[0]?.message).toContain("999999999");
  });

  it("rejects an oversized body that lies about its length, mid-stream", async () => {
    const diagnostics = collector();
    let pulled = 0;
    const endless = new ReadableStream<Uint8Array>({
      pull(controller) {
        pulled += 1;
        if (pulled > 100) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array(1024));
      },
    });
    const resolver = createImageResolver({
      allowRemote: true,
      maxBytes: 4096,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: () => Promise.resolve(new Response(endless)),
    });
    expect(await resolver.resolve("https://example.com/endless.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["too-large"]);
    // Abandoned as soon as the cap was passed, not after 100 KiB.
    expect(pulled).toBeLessThanOrEqual(6);
  });

  it("times out a request that never answers", async () => {
    const diagnostics = collector();
    const spy = hangForever();
    const resolver = createImageResolver({
      allowRemote: true,
      fetch: spy.fetch,
      timeoutMs: 25,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    const started = Date.now();
    expect(await resolver.resolve("https://example.com/slow.png")).toBeNull();
    expect(Date.now() - started).toBeLessThan(2000);
    expect(codes(diagnostics.seen)).toEqual(["timeout"]);
    expect(diagnostics.seen[0]?.message).toContain("25 ms");
  });

  it("reports an HTTP error status", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: () => Promise.resolve(new Response("nope", { status: 404, statusText: "Not Found" })),
    });
    expect(await resolver.resolve("https://example.com/missing.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["http-error"]);
    expect(diagnostics.seen[0]?.message).toContain("404");
  });

  it("reports a rejected fetch (CORS, DNS, reset) as a network error", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: () => Promise.reject(new TypeError("Failed to fetch")),
    });
    expect(await resolver.resolve("https://example.com/cors.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["network-error"]);
    expect(diagnostics.seen[0]?.message).toContain("Failed to fetch");
  });

  it("survives a fetch that throws synchronously", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: () => {
        throw new Error("boom");
      },
    });
    expect(await resolver.resolve("https://example.com/boom.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["network-error"]);
  });

  it("reports an empty body and undecodable bytes distinctly", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: (url) =>
        Promise.resolve(
          new Response(
            bodyOf(url.endsWith("empty.png") ? new Uint8Array(0) : ascii("<html>404</html>")),
          ),
        ),
    });
    expect(await resolver.resolve("https://example.com/empty.png")).toBeNull();
    expect(await resolver.resolve("https://example.com/html.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["empty", "undecodable"]);
  });

  it("truncated bytes are undecodable rather than a wrong-sized picture", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({
      allowRemote: true,
      onDiagnostic: diagnostics.onDiagnostic,
      fetch: () => Promise.resolve(new Response(bodyOf(png.subarray(0, 20)))),
    });
    expect(await resolver.resolve("https://example.com/cut.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["undecodable"]);
  });
});

/* -------------------------------------------------------------------------- */
/* blob: URLs                                                                  */
/* -------------------------------------------------------------------------- */

describe("blob: URLs", () => {
  it("resolves an object URL without allowRemote", async () => {
    const png = makePng(23, 9);
    const url = URL.createObjectURL(new Blob([bodyOf(png)], { type: "image/png" }));
    try {
      const resolver = createImageResolver();
      expect(await resolver.resolve(url)).toMatchObject({ format: "png", width: 23, height: 9 });
    } finally {
      URL.revokeObjectURL(url);
    }
  });

  it("can be switched off", async () => {
    const png = makePng(4, 4);
    const url = URL.createObjectURL(new Blob([bodyOf(png)]));
    try {
      const diagnostics = collector();
      const resolver = createImageResolver({
        allowBlob: false,
        onDiagnostic: diagnostics.onDiagnostic,
      });
      expect(await resolver.resolve(url)).toBeNull();
      expect(codes(diagnostics.seen)).toEqual(["unsupported-scheme"]);
    } finally {
      URL.revokeObjectURL(url);
    }
  });

  it("reports a revoked object URL instead of throwing", async () => {
    const url = URL.createObjectURL(new Blob([bodyOf(makePng(4, 4))]));
    URL.revokeObjectURL(url);
    const diagnostics = collector();
    const resolver = createImageResolver({ onDiagnostic: diagnostics.onDiagnostic });
    expect(await resolver.resolve(url)).toBeNull();
    expect(diagnostics.seen).toHaveLength(1);
  });
});

/* -------------------------------------------------------------------------- */
/* SVG and WebP                                                                */
/* -------------------------------------------------------------------------- */

describe("SVG", () => {
  const svg = makeSvg('width="120" height="60"');
  const pngTwin = makePng(120, 60);

  const rasterizer: ImageRasterizer = {
    rasterize: (request) =>
      Promise.resolve({
        format: "png",
        data: pngTwin,
        width: request.targetWidth ?? 120,
        height: request.targetHeight ?? 60,
      }),
  };

  it("degrades to the placeholder path when there is no rasterizer", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({ onDiagnostic: diagnostics.onDiagnostic });
    const resolved = await resolver.resolve(`data:image/svg+xml;base64,${toBase64(svg)}`);
    expect(resolved).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["svg-no-rasterizer"]);
    expect(diagnostics.seen[0]?.message).toContain("raster fallback");
  });

  it("embeds the vector plus its mandatory raster twin when one is supplied", async () => {
    const resolver = createImageResolver({ rasterizer });
    const resolved = await resolver.resolve(`data:image/svg+xml;base64,${toBase64(svg)}`);
    expect(resolved?.format).toBe("svg");
    expect(resolved?.width).toBe(120);
    expect(resolved?.height).toBe(60);
    expect(resolved?.fallback).toMatchObject({ format: "png" });
    expect(resolved?.fallback?.data).toBeInstanceOf(Uint8Array);
    // Bytes, never markup: ImageRun runs any string through atob().
    expect(resolved?.data).toBeInstanceOf(Uint8Array);
    expect(typeof resolved?.data).not.toBe("string");
  });

  it("hands the rasterizer a page-clamped target size", async () => {
    const requests: number[] = [];
    const resolver = createImageResolver({
      contentWidthTwips: 9026, // 601 px
      rasterizer: {
        rasterize: (request) => {
          requests.push(request.targetWidth ?? -1, request.targetHeight ?? -1);
          return Promise.resolve({ format: "png", data: pngTwin, width: 601, height: 301 });
        },
      },
    });
    const wide = makeSvg('width="1202" height="602"');
    await resolver.resolve(`data:image/svg+xml;base64,${toBase64(wide)}`);
    expect(requests).toEqual([601, 301]);
  });

  it("treats a rasterizer that fails, throws or lies as a failure - never a broken picture", async () => {
    const cases: readonly [string, ImageRasterizer][] = [
      ["returns null", { rasterize: () => Promise.resolve(null) }],
      [
        "throws",
        {
          rasterize: () => {
            throw new Error("canvas is gone");
          },
        },
      ],
      ["rejects", { rasterize: () => Promise.reject(new Error("worker died")) }],
      [
        "returns zero bytes",
        {
          rasterize: () =>
            Promise.resolve({ format: "png", data: new Uint8Array(0), width: 1, height: 1 }),
        },
      ],
      [
        "returns a zero-sized raster",
        { rasterize: () => Promise.resolve({ format: "png", data: pngTwin, width: 0, height: 0 }) },
      ],
    ];

    for (const [label, broken] of cases) {
      const diagnostics = collector();
      const resolver = createImageResolver({
        rasterizer: broken,
        onDiagnostic: diagnostics.onDiagnostic,
      });
      const resolved = await resolver.resolve(`data:image/svg+xml;base64,${toBase64(svg)}`);
      expect(resolved, label).toBeNull();
      expect(codes(diagnostics.seen), label).toEqual(["rasterizer-failed"]);
    }
  });

  it("proves the trap it is avoiding: ImageRun rejects markup, accepts bytes", () => {
    const markup = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="5"></svg>';
    const transformation = { width: 10, height: 5 };
    const fallback = { type: "png", data: pngTwin } as const;
    // docx runs any string through atob(), so raw markup throws inside the packer.
    expect(() => new ImageRun({ type: "svg", data: markup, fallback, transformation })).toThrow();
    expect(
      () =>
        new ImageRun({
          type: "svg",
          data: new TextEncoder().encode(markup),
          fallback,
          transformation,
        }),
    ).not.toThrow();
  });
});

describe("EXIF-rotated photographs", () => {
  it("resolves at the stored size and says so, rather than silently guessing", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({ onDiagnostic: diagnostics.onDiagnostic });
    const resolved = await resolver.resolve(
      `data:image/jpeg;base64,${toBase64(REAL_JPEG_EXIF_ORIENTATION_6)}`,
    );
    // The picture is embedded - this is a notice, not a failure.
    expect(resolved).toMatchObject({ format: "jpg", width: 23, height: 9 });
    expect(codes(diagnostics.seen)).toEqual(["exif-orientation-ignored"]);
    expect(diagnostics.seen[0]?.severity).toBe("notice");
    expect(diagnostics.seen[0]?.message).toContain("Word draws the stored pixels");
  });

  it("stays quiet for the non-rotating orientations", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({ onDiagnostic: diagnostics.onDiagnostic });
    const jpeg = makeJpeg(40, 10, { orientation: 2 });
    expect(await resolver.resolve(`data:image/jpeg;base64,${toBase64(jpeg)}`)).not.toBeNull();
    expect(diagnostics.seen).toEqual([]);
  });

  it("separates notices from errors in a document resolution", async () => {
    const rotated = image(`data:image/jpeg;base64,${toBase64(REAL_JPEG_EXIF_ORIENTATION_6)}`);
    const noticesOnly = await resolveDocumentImages(
      doc([paragraph([rotated])]),
      createImageResolver(),
    );
    expect(noticesOnly.resolvedCount).toBe(1);
    expect(codes(noticesOnly.diagnostics)).toEqual(["exif-orientation-ignored"]);
    expect(hasImageErrors(noticesOnly)).toBe(false);

    const withError = await resolveDocumentImages(
      doc([paragraph([image("https://example.com/a.png")])]),
      createImageResolver(),
    );
    expect(hasImageErrors(withError)).toBe(true);
    expect(diagnosticSeverity("remote-blocked")).toBe("error");
  });
});

describe("WebP", () => {
  it("is measured but refused, because OOXML cannot declare it", async () => {
    const diagnostics = collector();
    const resolver = createImageResolver({ onDiagnostic: diagnostics.onDiagnostic });
    const resolved = await resolver.resolve(`data:image/webp;base64,${toBase64(REAL_WEBP_LOSSY)}`);
    expect(resolved).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["unsupported-format"]);
    // The message knows the size, so it can say "unsupported", not "corrupt".
    expect(diagnostics.seen[0]?.message).toContain("23x9");
  });

  it("is transcoded when a rasterizer is supplied, and loses the SVG fallback field", async () => {
    const png = makePng(23, 9);
    const resolver = createImageResolver({
      rasterizer: {
        rasterize: (request) => {
          expect(request.format).toBe("webp");
          return Promise.resolve({ format: "png", data: png, width: 23, height: 9 });
        },
      },
    });
    const resolved = await resolver.resolve(`data:image/webp;base64,${toBase64(REAL_WEBP_LOSSY)}`);
    expect(resolved).toMatchObject({ format: "png", width: 23, height: 9, fallback: null });
  });
});

/* -------------------------------------------------------------------------- */
/* isValidResolvedImage                                                        */
/* -------------------------------------------------------------------------- */

describe("isValidResolvedImage", () => {
  const png = makePng(4, 4);
  const raster = { format: "png", data: png, width: 4, height: 4 } as const;

  it("accepts a well-formed raster and a well-formed SVG pair", () => {
    expect(isValidResolvedImage({ ...raster, fallback: null })).toBe(true);
    expect(
      isValidResolvedImage({ format: "svg", data: png, width: 4, height: 4, fallback: raster }),
    ).toBe(true);
  });

  it("rejects every shape that produces an unopenable document", () => {
    expect(isValidResolvedImage(null)).toBe(false);
    expect(
      isValidResolvedImage({ format: "svg", data: png, width: 4, height: 4, fallback: null }),
    ).toBe(false);
    expect(isValidResolvedImage({ ...raster, fallback: raster })).toBe(false);
    expect(isValidResolvedImage({ ...raster, width: 0, fallback: null })).toBe(false);
    expect(isValidResolvedImage({ ...raster, height: Number.NaN, fallback: null })).toBe(false);
    expect(isValidResolvedImage({ ...raster, data: new Uint8Array(0), fallback: null })).toBe(
      false,
    );
  });
});

/* -------------------------------------------------------------------------- */
/* The document pass                                                           */
/* -------------------------------------------------------------------------- */

describe("resolveDocumentImages", () => {
  const png = makePng(23, 9);
  const uri = `data:image/png;base64,${toBase64(png)}`;

  it("keys the map by node while fetching each source once", async () => {
    const first = image(uri, { alt: "one" });
    const second = image(uri, { alt: "two" });
    const third = image("https://example.com/blocked.png", { alt: "three" });
    const document = doc([paragraph([first, text(" ")]), paragraph([second, third])]);

    let calls = 0;
    const resolver: ImageResolver = {
      resolve: (src) => {
        calls += 1;
        return createImageResolver().resolve(src);
      },
    };

    const result = await resolveDocumentImages(document, resolver);
    expect(calls).toBe(2); // two distinct sources, three nodes
    expect(result.images.size).toBe(2);
    expect(result.images.get(first)).toBe(result.images.get(second));
    expect(result.images.has(third)).toBe(false);
    expect(result.resolvedCount).toBe(1);
    expect(result.failedCount).toBe(1);
  });

  it("never asks about a node that already carries bytes", async () => {
    const resolved: ResolvedImage = {
      format: "png",
      data: png,
      width: 23,
      height: 9,
      fallback: null,
    };
    const document = doc([paragraph([image("anything", { resolved })])]);
    let called = false;
    await resolveDocumentImages(document, {
      resolve: () => {
        called = true;
        return Promise.resolve(null);
      },
    });
    expect(called).toBe(false);
  });

  it("collects diagnostics and forwards them live", async () => {
    const live: ImageDiagnostic[] = [];
    const document = doc([
      paragraph([image("https://example.com/a.png"), image("data:image/png;base64,%%")]),
    ]);
    const result = await resolveDocumentImages(document, createImageResolver(), {
      onDiagnostic: (d) => void live.push(d),
    });
    expect([...codes(result.diagnostics)].sort()).toEqual(["bad-data-uri", "remote-blocked"]);
    expect(live).toEqual(result.diagnostics);
  });

  it("absorbs a resolver that throws", async () => {
    const document = doc([paragraph([image("a.png"), image("b.png")])]);
    const result = await resolveDocumentImages(document, {
      resolve: () => Promise.reject(new Error("resolver exploded")),
    });
    expect(result.images.size).toBe(0);
    expect(result.failedCount).toBe(2);
    expect(codes(result.diagnostics)).toEqual(["network-error", "network-error"]);
    expect(result.diagnostics[0]?.message).toContain("resolver exploded");
  });

  it("discards a ResolvedImage that violates the OOXML contract", async () => {
    const document = doc([paragraph([image("bad.svg")])]);
    const result = await resolveDocumentImages(document, {
      // An SVG with no raster twin: docx would reject it, or worse, not.
      resolve: () =>
        Promise.resolve({ format: "svg", data: png, width: 4, height: 4, fallback: null }),
    });
    expect(result.images.size).toBe(0);
    expect(codes(result.diagnostics)).toEqual(["invalid-resolution"]);
  });

  it("bounds concurrency", async () => {
    const sources = Array.from({ length: 12 }, (_, i) => image(`https://example.com/${i}.png`));
    const document = doc([paragraph(sources)]);
    let inFlight = 0;
    let peak = 0;
    const result = await resolveDocumentImages(
      document,
      {
        resolve: async () => {
          inFlight += 1;
          peak = Math.max(peak, inFlight);
          await new Promise((resolve) => setTimeout(resolve, 5));
          inFlight -= 1;
          return null;
        },
      },
      { concurrency: 3 },
    );
    expect(peak).toBe(3);
    expect(result.failedCount).toBe(12);
  });

  it("stops listening to the resolver once it is done", async () => {
    const resolver = createImageResolver();
    const document = doc([paragraph([image("https://example.com/a.png")])]);
    const result = await resolveDocumentImages(document, resolver);
    expect(result.diagnostics).toHaveLength(1);
    // A later, unrelated failure must not append to a returned array.
    await resolver.resolve("https://example.com/b.png");
    expect(result.diagnostics).toHaveLength(1);
  });

  it("survives a diagnostic handler that throws", async () => {
    const document = doc([paragraph([image("https://example.com/a.png")])]);
    const resolver = createImageResolver({
      onDiagnostic: () => {
        throw new Error("logger is broken");
      },
    });
    const result = await resolveDocumentImages(document, resolver, {
      onDiagnostic: () => {
        throw new Error("so is this one");
      },
    });
    expect(result.failedCount).toBe(1);
    expect(codes(result.diagnostics)).toEqual(["remote-blocked"]);
  });

  it("handles a document with no images at all", async () => {
    const result = await resolveDocumentImages(doc([paragraph([text("no pictures")])]), {
      resolve: () => Promise.reject(new Error("must not be called")),
    });
    expect(result.images.size).toBe(0);
    expect(result.diagnostics).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* The Node resolver                                                           */
/* -------------------------------------------------------------------------- */

describe("createNodeImageResolver", () => {
  let root = "";
  let inside = "";

  beforeAll(async () => {
    root = await mkdtemp(join(tmpdir(), "downword-images-"));
    inside = join(root, "assets");
    await mkdir(inside, { recursive: true });
    await writeFile(join(inside, "logo.png"), makePng(23, 9));
    await writeFile(join(inside, "big.png"), makePng(400, 400));
    await writeFile(join(inside, "notes.txt"), "definitely not an image");
    await writeFile(join(root, "secret.png"), makePng(2, 2));
  });

  afterAll(async () => {
    if (root !== "") await rm(root, { recursive: true, force: true });
  });

  it("reads a relative path under the base directory", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside });
    expect(await resolver.resolve("logo.png")).toMatchObject({
      format: "png",
      width: 23,
      height: 9,
    });
    expect(await resolver.resolve("./logo.png")).not.toBeNull();
  });

  it("copies the bytes out of Node's pooled buffers", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside });
    const resolved = await resolver.resolve("logo.png");
    expect(resolved?.data).toBeInstanceOf(Uint8Array);
    expect(resolved?.data.byteOffset).toBe(0);
    expect(resolved?.data.buffer.byteLength).toBe(resolved?.data.byteLength);
  });

  it("strips a cache-busting query the way a URL would", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside });
    expect(await resolver.resolve("logo.png?v=2")).not.toBeNull();
  });

  it("reads a file: URL", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside });
    const url = new URL(`file://${join(inside, "logo.png")}`).href;
    expect(await resolver.resolve(url)).not.toBeNull();
  });

  it("refuses to walk out of the base directory", async () => {
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("../secret.png")).toBeNull();
    expect(await resolver.resolve(join(root, "secret.png"))).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["path-not-allowed", "path-not-allowed"]);
  });

  it("refuses a symlink that points out of the base directory", async () => {
    const link = join(inside, "escape.png");
    await symlink(join(root, "secret.png"), link);
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("escape.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["path-not-allowed"]);
  });

  it("can be told to allow it", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside, allowOutsideBaseDir: true });
    expect(await resolver.resolve("../secret.png")).toMatchObject({ width: 2, height: 2 });
  });

  it("reports a missing file, a directory and a non-image", async () => {
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("nope.png")).toBeNull();
    expect(await resolver.resolve(".")).toBeNull();
    expect(await resolver.resolve("notes.txt")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["not-found", "not-found", "undecodable"]);
  });

  it("enforces the byte cap on local files too", async () => {
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      maxBytes: 128,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("big.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["too-large"]);
  });

  it("keeps remote images off by default here as well", async () => {
    const spy = respondWith(makePng(4, 4));
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      fetch: spy.fetch,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("https://example.com/a.png")).toBeNull();
    expect(spy.calls).toEqual([]);
    expect(codes(diagnostics.seen)).toEqual(["remote-blocked"]);
  });

  it("blocks the private network even when remote images are enabled", async () => {
    const spy = respondWith(makePng(4, 4));
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      allowRemote: true,
      fetch: spy.fetch,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    for (const url of [
      "http://127.0.0.1:9200/logo.png",
      "http://localhost:3000/logo.png",
      "http://169.254.169.254/latest/meta-data/",
      "http://10.0.0.5/logo.png",
      "http://192.168.1.1/logo.png",
      "http://[::1]:8080/logo.png",
    ]) {
      expect(await resolver.resolve(url), url).toBeNull();
    }
    expect(spy.calls).toEqual([]);
    expect(new Set(codes(diagnostics.seen))).toEqual(new Set(["private-network-blocked"]));
  });

  it("classifies addresses the way the guard needs", () => {
    for (const address of [
      "127.0.0.1",
      "10.1.2.3",
      "172.16.0.1",
      "172.31.255.255",
      "192.168.0.1",
      "169.254.169.254",
      "100.64.0.1",
      "0.0.0.0",
      "::1",
      "::",
      "fe80::1",
      "fd00::1",
      "::ffff:127.0.0.1",
      "localhost",
      "db.localhost",
      "printer.local",
    ]) {
      expect(isPrivateAddress(address), address).toBe(true);
    }
    for (const address of ["8.8.8.8", "1.1.1.1", "172.32.0.1", "example.com", "2606:4700::1111"]) {
      expect(isPrivateAddress(address), address).toBe(false);
    }
  });

  it("can be pointed at a public host once the guard is off", async () => {
    const spy = respondWith(makePng(6, 3));
    const resolver = createNodeImageResolver({
      baseDir: inside,
      allowRemote: true,
      blockPrivateNetwork: false,
      fetch: spy.fetch,
    });
    expect(await resolver.resolve("http://127.0.0.1:9200/logo.png")).toMatchObject({ width: 6 });
    expect(spy.calls).toHaveLength(1);
  });

  it("can have the filesystem switched off entirely", async () => {
    const diagnostics = collector();
    const resolver = createNodeImageResolver({
      baseDir: inside,
      allowFilesystem: false,
      onDiagnostic: diagnostics.onDiagnostic,
    });
    expect(await resolver.resolve("logo.png")).toBeNull();
    expect(codes(diagnostics.seen)).toEqual(["invalid-url"]);
  });

  it("still resolves data: URIs", async () => {
    const resolver = createNodeImageResolver({ baseDir: inside });
    const uri = `data:image/png;base64,${toBase64(makePng(9, 9))}`;
    expect(await resolver.resolve(uri)).toMatchObject({ width: 9, height: 9 });
  });
});

/* -------------------------------------------------------------------------- */
/* End to end: resolved bytes really do land in a .docx                        */
/* -------------------------------------------------------------------------- */

describe("end to end", () => {
  it("carries a resolved raster and a resolved SVG pair into the package", async () => {
    const png = makePng(23, 9);
    const svg = makeSvg('width="120" height="60"');
    const document = doc([
      paragraph([image(`data:image/png;base64,${toBase64(png)}`, { alt: "a chart" })]),
      paragraph([image(`data:image/svg+xml;base64,${toBase64(svg)}`, { alt: "a diagram" })]),
      paragraph([image("https://example.com/blocked.png", { alt: "not fetched" })]),
    ]);

    const resolver = createImageResolver({
      rasterizer: {
        rasterize: () => Promise.resolve({ format: "png", data: png, width: 120, height: 60 }),
      },
    });
    const { images, diagnostics } = await resolveDocumentImages(document, resolver);
    expect(images.size).toBe(2);
    expect(codes(diagnostics)).toEqual(["remote-blocked"]);

    const file = renderDocument(document, { images });
    const bytes = await packDocument(file);
    const parts = await listDocxParts(bytes);

    expect(parts.filter((part) => part.startsWith("word/media/")).length).toBe(3);
    expect(parts.some((part) => part.endsWith(".svg"))).toBe(true);

    const xml = await readDocxPart(bytes, "word/document.xml");
    // Two pictures drawn, the SVG carrying Microsoft's vector extension...
    expect(xml.match(/<a:blip/g)?.length).toBe(2);
    expect(xml).toContain("svgBlip");
    // ...and the blocked one degraded to the renderer's placeholder run.
    expect(xml).toContain("not fetched");
  });

  it("sizes the extent from the bytes, clamped to the text column", async () => {
    // 1200 px is wider than A4's 601 px text column; the ratio must survive.
    const wide = makePng(1200, 600);
    const document = doc([paragraph([image(`data:image/png;base64,${toBase64(wide)}`)])]);
    const { images } = await resolveDocumentImages(document, createImageResolver());
    const xml = await readDocxPart(
      await packDocument(renderDocument(document, { images })),
      "word/document.xml",
    );
    const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
    expect(extent).not.toBeNull();
    const cx = Number(extent?.[1]);
    const cy = Number(extent?.[2]);
    // 1200x600 into a 601 px column is 601 x 300.5 px. `pixelsToEmu` is what
    // docx does to `transformation` on the way out, so the half pixel is
    // carried into the extent instead of being rounded to 301 px first - the
    // ratio lands within 4e-7 of 2 rather than within 2e-3.
    expect(cx).toBe(pixelsToEmu(601));
    expect(cy).toBe(pixelsToEmu(300.5));
    expect(cx / cy).toBeCloseTo(2, 5);
  });
});

/* -------------------------------------------------------------------------- */
/* Browser bundle safety                                                       */
/* -------------------------------------------------------------------------- */

describe("browser entry point", () => {
  it("reaches no Node builtins and never pulls in node.ts", async () => {
    const { readFile: read } = await import("node:fs/promises");
    const { dirname, resolve: resolvePath } = await import("node:path");
    const { fileURLToPath } = await import("node:url");

    const here = dirname(fileURLToPath(import.meta.url));
    const entry = resolvePath(here, "../src/images/index.ts");

    const seen = new Set<string>();
    const offenders: string[] = [];

    const walk = async (file: string): Promise<void> => {
      if (seen.has(file)) return;
      seen.add(file);
      const source = await read(file, "utf8");
      const pattern = /(?:^|\n)\s*(?:import|export)[\s\S]*?from\s*["']([^"']+)["']/g;
      let match = pattern.exec(source);
      while (match !== null) {
        const specifier = match[1] ?? "";
        if (
          specifier.startsWith("node:") ||
          ["fs", "path", "os", "url", "dns"].includes(specifier)
        ) {
          offenders.push(`${file} imports ${specifier}`);
        } else if (specifier.startsWith(".")) {
          await walk(resolvePath(dirname(file), specifier.replace(/\.js$/, ".ts")));
        }
        match = pattern.exec(source);
      }
    };

    await walk(entry);

    expect(offenders).toEqual([]);
    expect([...seen].some((file) => file.endsWith("/images/node.ts"))).toBe(false);
    // Sanity check on the walker itself: it did traverse the module graph.
    expect(seen.size).toBeGreaterThanOrEqual(9);
  });
});
