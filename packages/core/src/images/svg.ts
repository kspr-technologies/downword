/**
 * SVG: detection and intrinsic sizing, without a DOM parser.
 *
 * ### The `ImageRun` SVG trap
 *
 * `docx`'s `ImageRun` accepts `type: "svg"`, which makes it look like SVG is a
 * first-class raster substitute. It is not. Two traps, both verified against
 * the shipped `docx@9.7.1`:
 *
 * 1. **A raster `fallback` is mandatory.** `SvgMediaOptions` requires it,
 *    because OOXML stores a vector picture as `<a:blip>` (the raster) *plus* an
 *    `asvg:svgBlip` extension pointing at the SVG. Word 2016+ draws the vector;
 *    every other consumer - LibreOffice, Google Docs, Pages, Word on older
 *    builds, the thumbnailer - draws the raster. Ship without one and the type
 *    system stops you; work around the type system and you ship a document that
 *    renders as a hole.
 * 2. **String data is base64-only.** `standardizeData()` treats any `string` as
 *    a data URI and runs it through `atob`, so passing raw `<svg>…</svg>`
 *    markup throws `InvalidCharacterError` from inside the packer. This module
 *    therefore only ever hands bytes (`Uint8Array`) downstream - never markup.
 *
 * So there is no way to embed an SVG "just as a vector". Either the caller
 * supplies an {@link import("./types.js").ImageRasterizer} to make the twin, or
 * the image degrades to the renderer's placeholder with an
 * `svg-no-rasterizer` diagnostic. Silently producing a broken picture is the
 * one outcome this module refuses.
 *
 * ### Sizing
 *
 * An SVG's intrinsic size comes from `width`/`height` on the root element, from
 * its `viewBox`, or - per CSS's rule for a replaced element with no intrinsic
 * dimensions - from the 300x150 default. Percentages are *not* an intrinsic
 * size (they are relative to a containing block that does not exist inside a
 * Word document), so they fall through to the `viewBox`.
 *
 * ### Security note
 *
 * The markup is embedded verbatim; it is neither parsed as XML nor sanitised.
 * Word's SVG renderer does not execute script, and the bytes never touch the
 * host page's DOM on this path - but a host that *also* previews the same
 * source in the browser is responsible for sanitising it there.
 */

import type { ImageProbe } from "./types.js";

/** CSS absolute length units, in pixels at 96 dpi. */
const UNIT_TO_PX: Readonly<Record<string, number>> = {
  "": 1,
  px: 1,
  pt: 96 / 72,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  q: 96 / 25.4 / 4,
};

/** CSS's intrinsic size for a replaced element that declares none. */
const CSS_DEFAULT_WIDTH = 300;
const CSS_DEFAULT_HEIGHT = 150;

/** How much of the file to decode when looking for the root element. */
const SNIFF_BYTES = 64 * 1024;

/** Decodes the head of a byte range as UTF-8, replacing anything invalid. */
function decodeHead(bytes: Uint8Array): string {
  const head = bytes.length > SNIFF_BYTES ? bytes.subarray(0, SNIFF_BYTES) : bytes;
  return new TextDecoder("utf-8", { fatal: false }).decode(head);
}

/**
 * Finds the root `<svg …>` start tag and returns its attribute text.
 *
 * Hand-scanned rather than regexed so that a `>` inside a quoted attribute
 * value (`<svg data-tip="a > b">`) does not truncate the tag.
 */
function findSvgTag(text: string): string | null {
  const match = /<svg(?=[\s/>])/i.exec(text);
  if (match === null) return null;
  const start = match.index + match[0].length;
  let quote: string | null = null;
  for (let i = start; i < text.length; i += 1) {
    const char = text.charAt(i);
    if (quote !== null) {
      if (char === quote) quote = null;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === ">") return text.slice(start, i);
  }
  return null;
}

const ATTRIBUTE_PATTERN = /([A-Za-z_:][-\w:.]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))/g;

/** Parses `name="value"` pairs out of a start tag's attribute text. */
function parseAttributes(tag: string): ReadonlyMap<string, string> {
  const attributes = new Map<string, string>();
  ATTRIBUTE_PATTERN.lastIndex = 0;
  let match = ATTRIBUTE_PATTERN.exec(tag);
  while (match !== null) {
    const name = (match[1] ?? "").toLowerCase();
    const value = match[2] ?? match[3] ?? match[4] ?? "";
    if (!attributes.has(name)) attributes.set(name, value);
    match = ATTRIBUTE_PATTERN.exec(tag);
  }
  return attributes;
}

/**
 * A CSS absolute length in pixels.
 *
 * Returns `null` for percentages and font-relative units: they describe a
 * fraction of something, and there is nothing to be a fraction of here.
 */
function parseLength(value: string | undefined): number | null {
  if (value === undefined) return null;
  const match = /^\s*([+-]?(?:\d+\.?\d*|\.\d+))(?:[eE][+-]?\d+)?\s*([a-zA-Z%]*)\s*$/.exec(value);
  if (match === null) return null;
  const magnitude = Number.parseFloat(match[1] ?? "");
  if (!Number.isFinite(magnitude) || magnitude <= 0) return null;
  const unit = (match[2] ?? "").toLowerCase();
  const factor = UNIT_TO_PX[unit];
  if (factor === undefined) return null;
  return magnitude * factor;
}

/** The `width height` half of a `viewBox`, when it is usable. */
function parseViewBox(value: string | undefined): { width: number; height: number } | null {
  if (value === undefined) return null;
  const parts = value.trim().split(/[\s,]+/);
  if (parts.length !== 4) return null;
  const width = Number.parseFloat(parts[2] ?? "");
  const height = Number.parseFloat(parts[3] ?? "");
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
  if (width <= 0 || height <= 0) return null;
  return { width, height };
}

/** Whether these bytes look like an SVG document. */
export function isSvg(bytes: Uint8Array): boolean {
  return findSvgTag(decodeHead(bytes)) !== null;
}

/**
 * Measures an SVG.
 *
 * @param bytes - The SVG source, UTF-8.
 * @returns A probe with `format: "svg"`, or `null` if there is no root element.
 *   `intrinsic` is `false` when the size is CSS's 300x150 default rather than
 *   something the file declared.
 */
export function probeSvg(bytes: Uint8Array): ImageProbe | null {
  const tag = findSvgTag(decodeHead(bytes));
  if (tag === null) return null;

  const attributes = parseAttributes(tag);
  const declaredWidth = parseLength(attributes.get("width"));
  const declaredHeight = parseLength(attributes.get("height"));
  const viewBox = parseViewBox(attributes.get("viewbox"));

  let width = declaredWidth;
  let height = declaredHeight;

  if (viewBox !== null) {
    const ratio = viewBox.width / viewBox.height;
    if (width === null && height === null) {
      width = viewBox.width;
      height = viewBox.height;
    } else if (width === null && height !== null) {
      width = height * ratio;
    } else if (height === null && width !== null) {
      height = width / ratio;
    }
  }

  const intrinsic = width !== null && height !== null;
  const finalWidth = width ?? CSS_DEFAULT_WIDTH;
  const finalHeight = height ?? CSS_DEFAULT_HEIGHT;

  return {
    format: "svg",
    width: Math.round(finalWidth),
    height: Math.round(finalHeight),
    orientation: null,
    orientedWidth: Math.round(finalWidth),
    orientedHeight: Math.round(finalHeight),
    intrinsic,
  };
}
