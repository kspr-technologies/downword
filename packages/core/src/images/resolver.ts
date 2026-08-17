/**
 * The {@link ImageResolver} implementation, and the egress switch.
 *
 * ### Default behaviour, stated plainly
 *
 * ```ts
 * createImageResolver();                       // data: and blob: only. No network. Ever.
 * createImageResolver({ allowRemote: true });  // http(s) too, capped at 10 MiB / 10 s.
 * NULL_IMAGE_RESOLVER;                         // nothing resolves at all.
 * ```
 *
 * `allowRemote` defaults to `false` and the check happens *before* the injected
 * `fetch` is touched, so a default resolver provably cannot make a request -
 * that is a property a test can assert (and does), not a promise in a comment.
 * `data:` and `blob:` stay on because neither leaves the process: a `data:` URI
 * carries its own bytes, and a `blob:` URL is a handle into this tab's memory.
 *
 * Anything that would open a socket is treated as remote, including a bare
 * relative path like `./diagram.png` - it resolves against the page's base URL
 * and hits the network exactly like an absolute one would. Grouping them stops
 * "it's just a local file" from becoming an accidental hole in the claim.
 *
 * ### Everything is a `null`
 *
 * `resolve()` has one job and cannot fail at it: unreachable, blocked,
 * oversized, timed out, CORS-refused, truncated, WebP, SVG-without-rasteriser,
 * or an outright bug in this file all come back as `null`, and the renderer
 * emits its placeholder. The whole body is wrapped, so even a programming error
 * here degrades one picture instead of failing the conversion.
 */

import type { ResolvedImage } from "../model.js";
import type { ImageResolver } from "../render/types.js";
import { parseDataUri, isDataUri } from "./data-uri.js";
import { decodeImage } from "./decode.js";
import { DEFAULT_CONTENT_WIDTH_TWIPS } from "./layout.js";
import { describeError, fetchImageBytes } from "./transport.js";
import {
  bytesFailure,
  diagnosticSeverity,
  resolveFetchPolicy,
  type DiagnosticSink,
  type ImageBytesResult,
  type ImageDiagnostic,
  type ImageDiagnosticHandler,
  type ImageFetchPolicy,
  type ImageRasterizer,
  type ResolvedFetchPolicy,
} from "./types.js";

/**
 * A pluggable byte source for schemes the core does not handle itself.
 *
 * Returning `null` means "not mine, keep looking", which is what lets the Node
 * resolver add filesystem support (`node.ts`) without the browser bundle ever
 * seeing `node:fs`.
 */
export type ImageLoader = (request: ImageLoadRequest) => Promise<ImageBytesResult | null>;

/** What a {@link ImageLoader} is asked to load. */
export interface ImageLoadRequest {
  /** The source exactly as authored. */
  readonly src: string;
  /** `src` parsed against the base URL, or `null` if it is not a URL. */
  readonly url: URL | null;
  /** Resolved caps and the egress switch. */
  readonly policy: ResolvedFetchPolicy;
}

/**
 * A last check before a remote request goes out.
 *
 * Returning a failure blocks the request; returning `null` allows it. The Node
 * resolver uses this for its SSRF guard - in a browser the same-origin policy
 * and CORS already do the job, and duplicating it would only break legitimate
 * intranet documents.
 */
export type RemoteGuard = (url: URL) => Promise<ImageBytesResult | null>;

/** Options for {@link createImageResolver}. */
export interface ImageResolverOptions extends ImageFetchPolicy {
  /**
   * Base for relative sources. Defaults to the document's base URI in a
   * browser, and to nothing anywhere else - where a relative path is
   * unresolvable and fails with `invalid-url` rather than guessing.
   */
  readonly baseUrl?: string | undefined;
  /** Resolve `data:` URIs. Defaults to `true`; no egress either way. */
  readonly allowData?: boolean | undefined;
  /** Resolve `blob:` object URLs. Defaults to `true`; reads this tab's memory, no egress. */
  readonly allowBlob?: boolean | undefined;
  /** Extra byte sources, tried after `data:` and before `blob:`/`http(s)`. */
  readonly loaders?: readonly ImageLoader[] | undefined;
  /** Vetoes a remote request after the URL is known. See {@link RemoteGuard}. */
  readonly guardRemote?: RemoteGuard | undefined;
  /** Converts SVG and WebP into embeddable rasters. See {@link ImageRasterizer}. */
  readonly rasterizer?: ImageRasterizer | undefined;
  /**
   * Usable text width in twips. Defaults to
   * {@link DEFAULT_CONTENT_WIDTH_TWIPS} (A4, one-inch margins).
   *
   * Its only effect is the target size suggested to an
   * {@link ImageRasterizer}; the binding clamp happens in the renderer at
   * layout time, where the real page geometry is known.
   */
  readonly contentWidthTwips?: number | undefined;
  /** Called once per problem. See {@link ImageDiagnosticHandler}. */
  readonly onDiagnostic?: ImageDiagnosticHandler | undefined;
  /**
   * Memoise by source, so the same URL referenced ten times is fetched once.
   * Defaults to `true`.
   *
   * Failures are cached too - a 404 becomes one diagnostic, not ten - which
   * also means a source resolved before an {@link ObservableImageResolver.observe}
   * listener was attached will not replay its diagnostic to that listener.
   */
  readonly cache?: boolean | undefined;
}

/** The document's base URI when there is a DOM, otherwise nothing. */
function ambientBaseUrl(): string | undefined {
  if (typeof document === "undefined") return undefined;
  const base: unknown = document.baseURI;
  return typeof base === "string" && base !== "" ? base : undefined;
}

/** Parses `src` against `base`, tolerating anything that is not a URL. */
function toUrl(src: string, base: string | undefined): URL | null {
  try {
    return base === undefined ? new URL(src) : new URL(src, base);
  } catch {
    return null;
  }
}

/**
 * A resolver that resolves nothing.
 *
 * The correct default for a host whose privacy story is "your document never
 * leaves the tab" and that does not need images from anywhere: images render as
 * the renderer's placeholder, and no code path exists that could fetch one.
 */
export const NULL_IMAGE_RESOLVER: ImageResolver = {
  resolve(): Promise<ResolvedImage | null> {
    return Promise.resolve(null);
  },
};

/**
 * An {@link ImageResolver} that will tell you *why* an image did not resolve.
 *
 * `ImageResolver` is a one-method interface owned by the renderer, and widening
 * it is not this module's call - but a resolver that swallows "the server
 * returned 403" is a support burden. So the extra channel is a structural
 * addition that {@link resolveDocumentImages} feature-detects: hook one up and
 * the pass's `diagnostics` array contains the transport-level reasons as well
 * as its own, wire up nothing and the behaviour is unchanged.
 */
export interface ObservableImageResolver extends ImageResolver {
  /** Registers a diagnostic listener. Call the returned function to remove it. */
  observe(handler: ImageDiagnosticHandler): () => void;
}

/** Whether a resolver can report diagnostics. See {@link ObservableImageResolver}. */
export function isObservableImageResolver(
  resolver: ImageResolver,
): resolver is ObservableImageResolver {
  return typeof (resolver as Partial<ObservableImageResolver>).observe === "function";
}

/**
 * Builds an {@link ImageResolver}.
 *
 * Browser-safe: this module and everything it imports are free of Node
 * builtins, which the test suite enforces by walking the import graph. Node
 * hosts should use `createNodeImageResolver` from `./node.js` instead - it adds
 * filesystem support and an SSRF guard on top of this.
 *
 * @param options - Caps, the egress switch and extension points. See {@link ImageResolverOptions}.
 * @returns A resolver whose `resolve()` never throws and never rejects.
 */
export function createImageResolver(options: ImageResolverOptions = {}): ObservableImageResolver {
  const policy = resolveFetchPolicy(options);
  const listeners = new Set<ImageDiagnosticHandler>();
  const sink: DiagnosticSink = {
    report(code, src, message) {
      if (options.onDiagnostic === undefined && listeners.size === 0) return;
      const diagnostic: ImageDiagnostic = {
        code,
        severity: diagnosticSeverity(code),
        src,
        message,
      };
      // Each listener is guarded on its own: one throwing logger must not
      // starve the others, and none of them may fail the conversion.
      for (const handler of [options.onDiagnostic, ...listeners]) {
        if (handler === undefined) continue;
        try {
          handler(diagnostic);
        } catch {
          /* a host whose logger throws is not this image's problem */
        }
      }
    },
  };
  const baseUrl = options.baseUrl ?? ambientBaseUrl();
  const allowData = options.allowData ?? true;
  const allowBlob = options.allowBlob ?? true;
  const loaders = options.loaders ?? [];
  const rasterizer = options.rasterizer ?? null;
  const contentWidthTwips = options.contentWidthTwips ?? DEFAULT_CONTENT_WIDTH_TWIPS;
  const cache = (options.cache ?? true) ? new Map<string, Promise<ResolvedImage | null>>() : null;

  async function load(src: string): Promise<ImageBytesResult> {
    if (isDataUri(src)) {
      if (!allowData) {
        return bytesFailure("unsupported-scheme", "data: URIs are disabled for this resolver");
      }
      return parseDataUri(src, policy.maxBytes);
    }

    const url = toUrl(src, baseUrl);

    for (const loader of loaders) {
      const result = await loader({ src, url, policy });
      if (result !== null) return result;
    }

    if (url === null) {
      return bytesFailure(
        "invalid-url",
        baseUrl === undefined
          ? `"${src}" is not an absolute URL and this resolver has no base URL to resolve it against`
          : `"${src}" could not be parsed as a URL`,
      );
    }

    const protocol = url.protocol.toLowerCase();

    if (protocol === "blob:") {
      if (!allowBlob) {
        return bytesFailure("unsupported-scheme", "blob: URLs are disabled for this resolver");
      }
      // An object URL is a handle into this process's memory; reading it is not
      // egress, so it is exempt from allowRemote. The caps still apply.
      return fetchImageBytes(url.href, policy);
    }

    if (protocol !== "http:" && protocol !== "https:") {
      return bytesFailure("unsupported-scheme", `this resolver cannot read ${protocol} sources`);
    }

    if (!policy.allowRemote) {
      // The gate. Nothing below this line runs for a default resolver.
      return bytesFailure(
        "remote-blocked",
        `remote images are disabled; pass { allowRemote: true } to let downword fetch ${url.origin}`,
      );
    }

    if (options.guardRemote !== undefined) {
      const blocked = await options.guardRemote(url);
      if (blocked !== null) return blocked;
    }

    return fetchImageBytes(url.href, policy);
  }

  async function resolveOnce(src: string): Promise<ResolvedImage | null> {
    try {
      const bytes = await load(src);
      if (!bytes.ok) {
        sink.report(bytes.code, src, `image "${src}": ${bytes.message}`);
        return null;
      }
      return await decodeImage(bytes.bytes, {
        src,
        sink,
        rasterizer,
        contentWidthTwips,
      });
    } catch (error: unknown) {
      // Belt and braces: a bug in here loses one image, not the document.
      sink.report("network-error", src, `image "${src}": ${describeError(error)}`);
      return null;
    }
  }

  return {
    resolve(src: string): Promise<ResolvedImage | null> {
      if (cache === null) return resolveOnce(src);
      const cached = cache.get(src);
      if (cached !== undefined) return cached;
      const pending = resolveOnce(src);
      cache.set(src, pending);
      return pending;
    },
    observe(handler: ImageDiagnosticHandler): () => void {
      listeners.add(handler);
      return () => {
        listeners.delete(handler);
      };
    },
  };
}
