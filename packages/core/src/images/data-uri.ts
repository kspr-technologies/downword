/**
 * `data:` URIs - the one image source that is always safe to resolve.
 *
 * A `data:` URI carries its bytes with it, so decoding one is not egress: no
 * socket is opened, nothing is disclosed, and the privacy claim in
 * {@link import("./types.js").ImageFetchPolicy} is untouched. That makes it the
 * default-on path, and in practice the common one - it is what a canvas,
 * `toDataURL()`, a Mermaid or chart plugin, and a pasted screenshot all produce.
 *
 * Grammar (RFC 2397):
 *
 * ```text
 *   data:[<mediatype>][;base64],<data>
 * ```
 *
 * Both halves are supported: `;base64` payloads and plain percent-encoded ones
 * (how an inline SVG usually travels). The byte cap applies here too - a 40 MiB
 * base64 blob pasted into a document is just as good a way to exhaust memory as
 * a 40 MiB download, and the string is length-checked *before* it is decoded so
 * the cap cannot be defeated by making us allocate first.
 */

import { decodeBase64, percentDecodeToBytes } from "./binary.js";
import { bytesFailure, type ImageBytesResult } from "./types.js";

/** Whether `src` is a `data:` URI. Cheap enough to call on every source. */
export function isDataUri(src: string): boolean {
  return /^data:/i.test(src.trimStart());
}

/**
 * Decodes a `data:` URI to bytes.
 *
 * @param src - The URI, exactly as authored.
 * @param maxBytes - Hard cap on the decoded payload.
 * @returns The bytes and the declared media type, or a typed failure. Never throws.
 */
export function parseDataUri(src: string, maxBytes: number): ImageBytesResult {
  const uri = src.trim();
  if (!isDataUri(uri)) return bytesFailure("bad-data-uri", "not a data: URI");

  const comma = uri.indexOf(",");
  if (comma === -1) {
    return bytesFailure("bad-data-uri", "data: URI has no comma separating header from payload");
  }

  const header = uri.slice(5, comma);
  const payload = uri.slice(comma + 1);

  const parameters = header.split(";");
  const mediaType = (parameters[0] ?? "").trim().toLowerCase();
  const base64 = parameters
    .slice(1)
    .some((parameter) => parameter.trim().toLowerCase() === "base64");

  // Reject before decoding. base64 inflates by 4/3 and percent-encoding by up
  // to 3x, so this bound is generous but still stops the pathological case of
  // allocating hundreds of megabytes to then find out it was too big.
  const inflation = base64 ? 2 : 4;
  if (payload.length > maxBytes * inflation) {
    return bytesFailure(
      "too-large",
      `data: URI payload is ~${payload.length} encoded characters, over the ${maxBytes}-byte cap`,
    );
  }

  const bytes = base64 ? decodeBase64(payload) : percentDecodeToBytes(payload);
  if (bytes === null) {
    return bytesFailure(
      "bad-data-uri",
      base64
        ? "data: URI payload is not valid base64"
        : "data: URI payload is not valid percent-encoding",
    );
  }
  if (bytes.length === 0) return bytesFailure("empty", "data: URI decoded to zero bytes");
  if (bytes.length > maxBytes) {
    return bytesFailure(
      "too-large",
      `data: URI decoded to ${bytes.length} bytes, over the ${maxBytes}-byte cap`,
    );
  }

  return { ok: true, bytes, mediaType: mediaType === "" ? null : mediaType };
}
