/**
 * Byte-level helpers shared by the sniffer and the `data:` URI parser.
 *
 * Everything here is dependency-free and runtime-agnostic on purpose: `Buffer`
 * exists only in Node, `atob` is a DOM API that rejects base64url and is
 * awkward about whitespace, and `decodeURIComponent` is *actively wrong* for
 * binary payloads because it insists the bytes are UTF-8. A `data:` URI holding
 * a percent-encoded PNG would come back mangled. Fifty lines of arithmetic buy
 * correctness in both runtimes and remove two footguns.
 */

/** Whether `bytes` contains `signature` at `offset`. */
export function hasSignature(bytes: Uint8Array, signature: readonly number[], offset = 0): boolean {
  if (offset < 0 || bytes.length < offset + signature.length) return false;
  for (let i = 0; i < signature.length; i += 1) {
    if (bytes[offset + i] !== signature[i]) return false;
  }
  return true;
}

/** Whether the ASCII text `text` appears at `offset`. */
export function hasAscii(bytes: Uint8Array, offset: number, text: string): boolean {
  if (offset < 0 || bytes.length < offset + text.length) return false;
  for (let i = 0; i < text.length; i += 1) {
    if (bytes[offset + i] !== text.charCodeAt(i)) return false;
  }
  return true;
}

/** Reads `length` ASCII bytes at `offset`, or `null` if they run past the end. */
export function readAscii(bytes: Uint8Array, offset: number, length: number): string | null {
  if (offset < 0 || bytes.length < offset + length) return null;
  let out = "";
  for (let i = 0; i < length; i += 1) {
    // Bounds are checked above, so the index is always populated.
    out += String.fromCharCode(bytes[offset + i] as number);
  }
  return out;
}

/** A `DataView` over exactly the bytes of `view`, honouring its byte offset. */
export function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

const BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/** `charCode -> 6-bit value`, `-1` for anything not in the alphabet. */
const BASE64_LOOKUP: Int8Array = (() => {
  const table = new Int8Array(256).fill(-1);
  for (let i = 0; i < BASE64_ALPHABET.length; i += 1) {
    table[BASE64_ALPHABET.charCodeAt(i)] = i;
  }
  // base64url, so a URL-safe payload pasted into a data: URI still decodes.
  table["-".charCodeAt(0)] = 62;
  table["_".charCodeAt(0)] = 63;
  return table;
})();

function isAsciiWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0c || code === 0x0d;
}

/**
 * Decodes base64 (or base64url) to bytes.
 *
 * Follows WHATWG "forgiving base64" in the two ways that matter for real
 * documents - ASCII whitespace is ignored (long data URIs get wrapped by
 * editors) and trailing `=` padding is optional - and is strict about
 * everything else. Returns `null` rather than throwing on bad input.
 */
export function decodeBase64(input: string): Uint8Array | null {
  // Worst case output size: every char is a real 6-bit digit.
  const out = new Uint8Array(Math.ceil((input.length * 3) / 4));
  let length = 0;
  let accumulator = 0;
  let bits = 0;
  let padding = 0;

  for (let i = 0; i < input.length; i += 1) {
    const code = input.charCodeAt(i);
    if (isAsciiWhitespace(code)) continue;
    if (code === 0x3d /* = */) {
      padding += 1;
      if (padding > 2) return null;
      continue;
    }
    // Padding must be the last thing in the string.
    if (padding > 0) return null;
    if (code > 0xff) return null;
    const value = BASE64_LOOKUP[code] as number;
    if (value < 0) return null;
    accumulator = (accumulator << 6) | value;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out[length] = (accumulator >> bits) & 0xff;
      length += 1;
    }
  }

  // A single leftover 6-bit group cannot be part of any byte: the input was
  // truncated mid-character.
  if (bits >= 6) return null;
  // Whatever bits remain must be zero padding, not dropped data.
  if (bits > 0 && (accumulator & ((1 << bits) - 1)) !== 0) return null;

  // `slice`, not `subarray`: the bytes end up in `ImageRun`, and handing a
  // consumer a view onto a larger buffer invites `.buffer`-based code to read
  // the slack.
  return length === out.length ? out : out.slice(0, length);
}

/**
 * Percent-decodes to *bytes*, the way RFC 3986 actually defines it.
 *
 * `decodeURIComponent` would UTF-8-decode the result into a string, which
 * silently corrupts every non-text payload; this walks the octets instead.
 * Literal characters outside `%XX` escapes are encoded as UTF-8, which is what
 * a browser does with a `data:image/svg+xml,<svg …>` URI.
 */
export function percentDecodeToBytes(input: string): Uint8Array | null {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  let total = 0;
  let index = 0;

  const push = (chunk: Uint8Array): void => {
    if (chunk.length === 0) return;
    chunks.push(chunk);
    total += chunk.length;
  };

  while (index < input.length) {
    const next = input.indexOf("%", index);
    if (next === -1) {
      push(encoder.encode(input.slice(index)));
      break;
    }
    if (next > index) push(encoder.encode(input.slice(index, next)));
    const hex = input.slice(next + 1, next + 3);
    if (hex.length < 2 || !/^[0-9a-fA-F]{2}$/.test(hex)) return null;
    push(Uint8Array.of(Number.parseInt(hex, 16)));
    index = next + 3;
  }

  return concatBytes(chunks, total);
}

/** Joins byte chunks into one buffer. `total` must be their combined length. */
export function concatBytes(chunks: readonly Uint8Array[], total: number): Uint8Array {
  if (chunks.length === 1 && (chunks[0] as Uint8Array).length === total) {
    return chunks[0] as Uint8Array;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
