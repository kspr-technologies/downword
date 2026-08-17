/**
 * Images: resolution, sniffing and sizing.
 *
 * The browser-safe entry point for the image pipeline. Everything reachable
 * from here is free of Node builtins and of the DOM, so it runs unchanged in a
 * tab, a worker, Node and a server handler; the one module that is not,
 * `./node.js`, is deliberately absent from this barrel so that a bundler never
 * has to resolve `node:fs`.
 *
 * ```ts
 * import {
 *   createImageResolver,
 *   resolveDocumentImages,
 *   NULL_IMAGE_RESOLVER,
 * } from "./images/index.js";
 *
 * // Default: data: and blob: only. No network request is possible.
 * const resolver = createImageResolver();
 *
 * // Opt in, explicitly, to egress - capped at 10 MiB and 10 s per image.
 * const online = createImageResolver({ allowRemote: true });
 *
 * const { images, diagnostics } = await resolveDocumentImages(doc, resolver);
 * const file = renderDocument(doc, { images });
 * ```
 *
 * Failure is never an exception. An image that cannot be fetched, is too big,
 * times out, is in a format OOXML cannot hold, or is simply corrupt resolves to
 * `null`, and the renderer emits its placeholder run in place of the picture;
 * the reason arrives as an {@link ImageDiagnostic} rather than as a thrown
 * error, so one bad URL can never fail a conversion.
 */

export {
  concatBytes,
  decodeBase64,
  hasAscii,
  hasSignature,
  percentDecodeToBytes,
  readAscii,
} from "./binary.js";

export { isDataUri, parseDataUri } from "./data-uri.js";

export { decodeImage, isValidResolvedImage, type DecodeImageOptions } from "./decode.js";

export {
  DEFAULT_CONTENT_WIDTH_TWIPS,
  EMU_PER_INCH,
  EMU_PER_PIXEL,
  PIXELS_PER_INCH,
  TWIPS_PER_PIXEL,
  fitToWidth,
  pixelsToEmu,
  pixelsToTwips,
  twipsToPixels,
  type PixelSize,
} from "./layout.js";

export {
  hasImageErrors,
  resolveDocumentImages,
  type DocumentImageResolution,
  type ResolveDocumentImagesOptions,
} from "./pass.js";

export { probeImage } from "./probe.js";

export {
  NULL_IMAGE_RESOLVER,
  createImageResolver,
  isObservableImageResolver,
  type ImageLoadRequest,
  type ImageLoader,
  type ImageResolverOptions,
  type ObservableImageResolver,
  type RemoteGuard,
} from "./resolver.js";

export { isSvg, probeSvg } from "./svg.js";

export { fetchImageBytes } from "./transport.js";

export {
  DEFAULT_MAX_BYTES,
  DEFAULT_TIMEOUT_MS,
  createDiagnosticSink,
  diagnosticSeverity,
  resolveFetchPolicy,
  type DiagnosticSink,
  type FetchLike,
  type ImageBytesResult,
  type ImageDiagnostic,
  type ImageDiagnosticCode,
  type ImageDiagnosticHandler,
  type ImageDiagnosticSeverity,
  type ImageFetchPolicy,
  type ImageProbe,
  type ImageRasterizer,
  type ProbedImageFormat,
  type RasterizeRequest,
  type ResolvedFetchPolicy,
} from "./types.js";
