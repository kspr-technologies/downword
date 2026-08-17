/**
 * Shared vocabulary for the image pipeline.
 *
 * The pipeline has four stages, and each one is a separate module so that a
 * bundler can drop the parts a given host does not use:
 *
 * ```text
 *   src ──▶ transport ──▶ bytes ──▶ probe/decode ──▶ ResolvedImage ──▶ renderer
 *           (data:/blob:/http(s)/fs)   (headers only, no DOM)
 * ```
 *
 * Two invariants hold everywhere in this directory:
 *
 * 1. **Nothing throws.** Every failure is a `null` plus an
 *    {@link ImageDiagnostic}. `convert()` must never reject because a picture
 *    404'd, and the renderer already has a placeholder path for a missing
 *    image, so an exception would only ever be a worse version of `null`.
 * 2. **No egress unless the caller asked for it.** The hosted site claims your
 *    document never leaves the tab; that claim is only true if the default
 *    resolver cannot make a network request. See {@link ImageFetchPolicy}.
 */

import type { RasterImageFormat, ResolvedRasterImage } from "../model.js";
import type { BaseWarning, WarningSeverity } from "../warnings.js";

/* -------------------------------------------------------------------------- */
/* Probing                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * Every format {@link import("./probe.js").probeImage} recognises.
 *
 * A superset of {@link RasterImageFormat}: `webp` and `svg` can be *measured*
 * but not embedded directly (see {@link ImageRasterizer}), and knowing their
 * size is what lets us tell "unsupported format" apart from "corrupt bytes".
 */
export type ProbedImageFormat = RasterImageFormat | "webp" | "svg";

/** What the bytes say about themselves, read from headers alone. */
export interface ImageProbe {
  readonly format: ProbedImageFormat;
  /** Stored raster width in pixels — the pixels a decoder will actually emit. */
  readonly width: number;
  /** Stored raster height in pixels. */
  readonly height: number;
  /**
   * JPEG EXIF orientation (1-8), or `null` when absent or not a JPEG.
   *
   * See {@link ImageProbe.orientedWidth} for what we do (and deliberately do
   * not do) with it.
   */
  readonly orientation: number | null;
  /**
   * Width *after* applying EXIF orientation — swapped with the height for the
   * four rotated orientations (5-8), equal to {@link ImageProbe.width}
   * otherwise.
   *
   * This is reported but **not** used to size the picture in the document; see
   * the note on EXIF in `probe.ts` for why.
   */
  readonly orientedWidth: number;
  /** Height after applying EXIF orientation. See {@link ImageProbe.orientedWidth}. */
  readonly orientedHeight: number;
  /**
   * `true` when the dimensions were read out of the file.
   *
   * Only ever `false` for an SVG that declares neither `width`/`height` nor a
   * `viewBox`, where CSS says a replaced element defaults to 300x150.
   */
  readonly intrinsic: boolean;
}

/* -------------------------------------------------------------------------- */
/* Diagnostics                                                                 */
/* -------------------------------------------------------------------------- */

/** Machine-readable reason an image did or did not resolve. */
export type ImageDiagnosticCode =
  /** The `data:` URI was malformed, or its payload was not valid base64. */
  | "bad-data-uri"
  /** A scheme this resolver does not handle (`file:` in a browser, `ftp:`, …). */
  | "unsupported-scheme"
  /** `new URL()` rejected the source, and no base URL made it absolute. */
  | "invalid-url"
  /** A network fetch was required but `allowRemote` is `false`. This is the default. */
  | "remote-blocked"
  /** The host resolves to loopback / link-local / RFC1918 space (SSRF guard). */
  | "private-network-blocked"
  /** A filesystem path escaped the resolver's `baseDir`. */
  | "path-not-allowed"
  /** The file does not exist, or is not a regular file. */
  | "not-found"
  /** The server answered, but not with 2xx. */
  | "http-error"
  /** DNS failure, connection reset, CORS rejection - anything `fetch()` rejects with. */
  | "network-error"
  /** The request exceeded `timeoutMs` and was aborted. */
  | "timeout"
  /** The payload exceeded `maxBytes`, declared or streamed. */
  | "too-large"
  /** Zero bytes. */
  | "empty"
  /** Bytes arrived but no decoder recognised them. */
  | "undecodable"
  /** A real image in a format OOXML cannot embed (WebP, AVIF, …) and no rasteriser was supplied. */
  | "unsupported-format"
  /** An SVG with no rasteriser to build the mandatory raster twin. */
  | "svg-no-rasterizer"
  /** The supplied {@link ImageRasterizer} returned `null` or something unusable. */
  | "rasterizer-failed"
  /** A third-party {@link import("../render/types.js").ImageResolver} broke the {@link import("../model.js").ResolvedImage} contract. */
  | "invalid-resolution"
  /** Informational: the JPEG carries a rotating EXIF orientation, which Word ignores. */
  | "exif-orientation-ignored";

/**
 * How much a diagnostic matters. The package-wide {@link WarningSeverity}.
 *
 * `"error"` means the image did not resolve and the renderer will emit a
 * placeholder. `"notice"` means it resolved, but with a caveat worth surfacing.
 */
export type ImageDiagnosticSeverity = WarningSeverity;

/** A non-fatal problem encountered while resolving one image. See {@link BaseWarning}. */
export interface ImageDiagnostic extends BaseWarning {
  readonly code: ImageDiagnosticCode;
  /** The source exactly as authored, so the message can be traced back to the markdown. */
  readonly src: string;
}

/** Receives one {@link ImageDiagnostic} per problem. Must not throw (we catch it if it does). */
export type ImageDiagnosticHandler = (diagnostic: ImageDiagnostic) => void;

const DIAGNOSTIC_SEVERITY: Readonly<Record<ImageDiagnosticCode, ImageDiagnosticSeverity>> = {
  "bad-data-uri": "error",
  "unsupported-scheme": "error",
  "invalid-url": "error",
  "remote-blocked": "error",
  "private-network-blocked": "error",
  "path-not-allowed": "error",
  "not-found": "error",
  "http-error": "error",
  "network-error": "error",
  timeout: "error",
  "too-large": "error",
  empty: "error",
  undecodable: "error",
  "unsupported-format": "error",
  "svg-no-rasterizer": "error",
  "rasterizer-failed": "error",
  "invalid-resolution": "error",
  "exif-orientation-ignored": "notice",
};

/** The severity of a diagnostic code. Total by construction. */
export function diagnosticSeverity(code: ImageDiagnosticCode): ImageDiagnosticSeverity {
  return DIAGNOSTIC_SEVERITY[code];
}

/** Somewhere to send diagnostics that swallows handler exceptions. */
export interface DiagnosticSink {
  report(code: ImageDiagnosticCode, src: string, message: string): void;
}

/** A sink that forwards to `handler`, or drops everything when it is absent. */
export function createDiagnosticSink(
  handler?: ImageDiagnosticHandler | null | undefined,
): DiagnosticSink {
  if (handler === null || handler === undefined) {
    return {
      report() {
        /* no handler: diagnostics are still cheap to produce, so callers need no guard */
      },
    };
  }
  return {
    report(code, src, message) {
      try {
        handler({ code, severity: diagnosticSeverity(code), src, message });
      } catch {
        // A host whose logger throws must not take the conversion down with it.
      }
    },
  };
}

/* -------------------------------------------------------------------------- */
/* Byte results                                                                */
/* -------------------------------------------------------------------------- */

/**
 * The outcome of trying to obtain bytes for a source.
 *
 * A result union rather than exceptions: every transport (data URI, fetch,
 * filesystem) fails in the same six or seven ways, and a union makes the
 * caller enumerate them instead of pattern-matching on error messages.
 */
export type ImageBytesResult =
  | {
      readonly ok: true;
      readonly bytes: Uint8Array;
      /** Declared media type, when the transport supplied one. A hint only - the sniffer decides. */
      readonly mediaType: string | null;
    }
  | {
      readonly ok: false;
      readonly code: ImageDiagnosticCode;
      readonly message: string;
    };

/** Convenience constructor for a failed {@link ImageBytesResult}. */
export function bytesFailure(code: ImageDiagnosticCode, message: string): ImageBytesResult {
  return { ok: false, code, message };
}

/* -------------------------------------------------------------------------- */
/* Rasterising the formats OOXML cannot hold                                   */
/* -------------------------------------------------------------------------- */

/** A request to turn a non-embeddable image into one `ImageRun` accepts. */
export interface RasterizeRequest {
  /** The original bytes: SVG markup (UTF-8) or a WebP file. */
  readonly data: Uint8Array;
  readonly format: "svg" | "webp";
  /** The source as authored, for error messages. */
  readonly src: string;
  /** Intrinsic size in CSS pixels, or `null` when the file declares none. */
  readonly intrinsicWidth: number | null;
  readonly intrinsicHeight: number | null;
  /**
   * The size the picture will occupy in the document, in CSS pixels, once
   * clamped to the page's content width - the sensible raster target. `null`
   * when the caller did not say how wide the page is.
   */
  readonly targetWidth: number | null;
  readonly targetHeight: number | null;
}

/**
 * Turns an SVG or WebP into a raster OOXML can embed.
 *
 * Deliberately an injected interface rather than an implementation: rasterising
 * needs a canvas (browser), a headless browser, or `sharp`/`resvg` (Node), and
 * none of those belong in a dependency-light core. Without one:
 *
 * - **SVG** cannot be embedded at all. OOXML stores an SVG *plus* a raster twin
 *   (`asvg:svgBlip` + `a:blip`), and `docx`'s `ImageRun` makes the `fallback`
 *   mandatory for `type: "svg"`, so there is no legal half-measure. The image
 *   degrades to the renderer's placeholder and a `svg-no-rasterizer` diagnostic.
 * - **WebP** cannot be embedded either: `ImageRun` accepts only
 *   `png | jpg | gif | bmp`. It degrades with `unsupported-format`.
 *
 * A returned raster is validated before use - a rasteriser that hands back zero
 * bytes or nonsense dimensions is treated as a failure, not embedded.
 */
export interface ImageRasterizer {
  rasterize(request: RasterizeRequest): Promise<ResolvedRasterImage | null>;
}

/* -------------------------------------------------------------------------- */
/* Fetch policy                                                                */
/* -------------------------------------------------------------------------- */

/** The subset of `fetch` this package uses. Injectable so tests need no network. */
export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

/** 10 MiB. Roughly a 5000x3000 photo; far more than a document needs. */
export const DEFAULT_MAX_BYTES = 10 * 1024 * 1024;

/** 10 seconds, wall clock, covering connect *and* body. */
export const DEFAULT_TIMEOUT_MS = 10_000;

/**
 * Caps and the egress switch.
 *
 * ### `allowRemote` defaults to `false`, and that is load-bearing
 *
 * downword's pitch is that your document never leaves the tab. An image
 * resolver is the one component that could quietly break that promise: a
 * markdown file pasted from a chat window frequently contains
 * `![](https://tracker.example/pixel.png)`, and fetching it would leak both the
 * fact that you opened the document and your IP to a third party.
 *
 * So the default is no egress at all. `data:` and `blob:` URLs still resolve -
 * neither leaves the process - but anything that would open a socket returns
 * `null` with a `remote-blocked` diagnostic, and the injected `fetch` is not
 * even called. Turning it on is a one-word, greppable opt-in:
 *
 * ```ts
 * createImageResolver({ allowRemote: true });
 * ```
 *
 * There is intentionally **no host allowlist**: an allowlist implies the
 * unlisted hosts are the dangerous ones, when the real risk is the request
 * existing at all. What there is instead is a hard byte cap and a hard timeout,
 * so an enabled resolver still cannot be turned into a denial-of-service.
 */
export interface ImageFetchPolicy {
  /**
   * Allow network requests. **Defaults to `false`.** With it off, `http:` and
   * `https:` sources - and relative paths, which would resolve to one - fail
   * with `remote-blocked` before any request is made.
   */
  readonly allowRemote?: boolean | undefined;
  /** Hard byte cap, declared or streamed. Defaults to {@link DEFAULT_MAX_BYTES} (10 MiB). */
  readonly maxBytes?: number | undefined;
  /** Abort after this many milliseconds. Defaults to {@link DEFAULT_TIMEOUT_MS} (10 s). */
  readonly timeoutMs?: number | undefined;
  /** `fetch` implementation. Defaults to the global one; injected in tests. */
  readonly fetch?: FetchLike | undefined;
  /** Aborts the whole resolution, on top of the per-request timeout. */
  readonly signal?: AbortSignal | undefined;
}

/** An {@link ImageFetchPolicy} with every default filled in. */
export interface ResolvedFetchPolicy {
  readonly allowRemote: boolean;
  readonly maxBytes: number;
  readonly timeoutMs: number;
  readonly fetch: FetchLike | null;
  readonly signal: AbortSignal | null;
}

/**
 * Fills in the policy defaults.
 *
 * `fetch` resolves to `null` rather than throwing when the runtime has none:
 * a missing `fetch` then reads as a `network-error` on the one image that
 * needed it, instead of a module-load crash for documents that have no images.
 */
export function resolveFetchPolicy(policy: ImageFetchPolicy = {}): ResolvedFetchPolicy {
  const globalFetch: FetchLike | null =
    typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : null;
  const maxBytes = policy.maxBytes;
  const timeoutMs = policy.timeoutMs;
  return {
    allowRemote: policy.allowRemote ?? false,
    maxBytes:
      maxBytes !== undefined && Number.isFinite(maxBytes) && maxBytes > 0
        ? Math.floor(maxBytes)
        : DEFAULT_MAX_BYTES,
    timeoutMs:
      timeoutMs !== undefined && Number.isFinite(timeoutMs) && timeoutMs > 0
        ? Math.floor(timeoutMs)
        : DEFAULT_TIMEOUT_MS,
    fetch: policy.fetch ?? globalFetch,
    signal: policy.signal ?? null,
  };
}
