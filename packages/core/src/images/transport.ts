/**
 * The one place this package is allowed to open a socket.
 *
 * Everything about it is defensive, because an image URL in a markdown document
 * is attacker-controlled input in the general case: it arrives in a file the
 * user pasted from somewhere else.
 *
 * - **Capped.** The body is read incrementally and abandoned the moment it
 *   passes `maxBytes`. A declared `Content-Length` short-circuits that, but a
 *   lying or absent one changes nothing: the stream is counted as it arrives.
 *   An endpoint that streams `/dev/urandom` forever costs us 10 MiB, once.
 * - **Timed out.** One `AbortController` covers connect *and* body, so a server
 *   that accepts the connection and then dribbles a byte a minute cannot pin a
 *   conversion open. `fetch`'s own timeout would not do this - it does not have
 *   one.
 * - **Cookie-free.** `credentials: "omit"` and `referrer: "no-referrer"`: an
 *   image fetch must not carry the user's session to a third party, and must
 *   not tell that third party which page asked.
 * - **Non-throwing.** Every rejection - DNS, CORS, reset, abort - comes back as
 *   a typed {@link ImageBytesResult}. Nothing here can produce an unhandled
 *   rejection: the fetch promise, the reader, and the cancel are each awaited
 *   inside a `try`, and the timer is cleared in a `finally`.
 *
 * Note what is *not* here: the decision to fetch at all. That lives in
 * `resolver.ts`, gated on `allowRemote`, so this module is never reached for a
 * default-configured resolver.
 */

import { concatBytes } from "./binary.js";
import { bytesFailure, type ImageBytesResult, type ResolvedFetchPolicy } from "./types.js";

/** Reads a `Content-Length` that is present, numeric and sane. */
function declaredLength(response: Response): number | null {
  const header = response.headers.get("content-length");
  if (header === null) return null;
  const value = Number.parseInt(header, 10);
  return Number.isFinite(value) && value >= 0 ? value : null;
}

/** The media type from `Content-Type`, without its parameters. */
function declaredMediaType(response: Response): string | null {
  const header = response.headers.get("content-type");
  if (header === null) return null;
  const type = (header.split(";")[0] ?? "").trim().toLowerCase();
  return type === "" ? null : type;
}

/**
 * Drains a response body, giving up as soon as it exceeds the cap.
 *
 * Falls back to `arrayBuffer()` when the runtime hands back a body-less
 * `Response` (an old polyfill, or a stubbed one in a test), and re-checks the
 * size afterwards so the cap holds on that path too.
 */
async function readCapped(
  response: Response,
  maxBytes: number,
): Promise<{ ok: true; bytes: Uint8Array } | { ok: false; message: string }> {
  const body = response.body;
  if (body === null || typeof body.getReader !== "function") {
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > maxBytes) {
      return { ok: false, message: `response body is ${buffer.byteLength} bytes` };
    }
    return { ok: true, bytes: new Uint8Array(buffer) };
  }

  const reader = body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value === undefined) continue;
      total += value.byteLength;
      if (total > maxBytes) {
        return { ok: false, message: `response body exceeded ${maxBytes} bytes` };
      }
      chunks.push(value);
    }
  } finally {
    // Releases the connection whether we finished or bailed out. The rejection
    // `cancel()` can produce on an already-errored stream is deliberately
    // swallowed - the outcome is decided above.
    await reader.cancel().catch(() => undefined);
  }
  return { ok: true, bytes: concatBytes(chunks, total) };
}

/**
 * Fetches image bytes under {@link ResolvedFetchPolicy}'s caps.
 *
 * @param url - An absolute URL. The caller has already decided this is allowed.
 * @param policy - Resolved caps, timeout and `fetch` implementation.
 * @returns Bytes plus the declared media type, or a typed failure. Never throws.
 */
export async function fetchImageBytes(
  url: string,
  policy: ResolvedFetchPolicy,
): Promise<ImageBytesResult> {
  const fetchImpl = policy.fetch;
  if (fetchImpl === null) {
    return bytesFailure("network-error", "no fetch implementation is available in this runtime");
  }

  const controller = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, policy.timeoutMs);

  const external = policy.signal;
  const onExternalAbort = (): void => {
    controller.abort();
  };
  if (external !== null) {
    if (external.aborted) controller.abort();
    else external.addEventListener("abort", onExternalAbort, { once: true });
  }

  try {
    const response = await fetchImpl(url, {
      signal: controller.signal,
      credentials: "omit",
      referrerPolicy: "no-referrer",
      redirect: "follow",
    });

    if (!response.ok) {
      return bytesFailure(
        "http-error",
        `server responded ${response.status} ${response.statusText}`,
      );
    }

    const length = declaredLength(response);
    if (length !== null && length > policy.maxBytes) {
      // Declared too big: hang up before reading a single byte of it.
      controller.abort();
      return bytesFailure(
        "too-large",
        `server declared ${length} bytes, over the ${policy.maxBytes}-byte cap`,
      );
    }

    const read = await readCapped(response, policy.maxBytes);
    if (!read.ok) {
      controller.abort();
      return bytesFailure("too-large", read.message);
    }
    if (read.bytes.length === 0) return bytesFailure("empty", "server returned zero bytes");

    return { ok: true, bytes: read.bytes, mediaType: declaredMediaType(response) };
  } catch (error: unknown) {
    if (timedOut) {
      return bytesFailure("timeout", `request timed out after ${policy.timeoutMs} ms`);
    }
    if (external !== null && external.aborted) {
      return bytesFailure("network-error", "request was aborted by the caller");
    }
    return bytesFailure("network-error", describeError(error));
  } finally {
    clearTimeout(timer);
    if (external !== null) external.removeEventListener("abort", onExternalAbort);
  }
}

/**
 * A short, safe description of whatever `fetch` rejected with.
 *
 * `fetch` rejects with a deliberately vague `TypeError` for a CORS failure, so
 * "failed to fetch" is frequently all there is to say - the browser withholds
 * the real reason on purpose.
 */
export function describeError(error: unknown): string {
  if (error instanceof Error) {
    const cause = error.cause;
    if (cause instanceof Error && cause.message !== "") {
      return `${error.message} (${cause.message})`;
    }
    return error.message === "" ? error.name : error.message;
  }
  return typeof error === "string" ? error : "unknown error";
}
