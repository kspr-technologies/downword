/**
 * The resolve pass: walk the document, fetch every picture, hand the renderer a
 * map it can read synchronously.
 *
 * This is the seam between the async world (network, filesystem, canvas) and
 * `renderDocument()`, which never awaits anything. The pass is a superset of
 * `render/prepare.ts`'s `prepareImages()`: same identity-keyed
 * {@link ImageMap} output, plus bounded concurrency, plus collected
 * diagnostics, plus validation of whatever a third-party resolver hands back.
 *
 * ```ts
 * const { images, diagnostics } = await resolveDocumentImages(doc, resolver);
 * const file = renderDocument(doc, { images });
 * ```
 *
 * Three properties worth stating, because each one is a bug someone will
 * otherwise hit:
 *
 * - **Identity-keyed, source-memoised.** The map is keyed by node, so two
 *   `![](logo.png)` references get their own entries; the *work* is keyed by
 *   `src`, so the bytes are fetched once. (`docx` then dedupes the media part
 *   again by content hash when packing.)
 * - **Bounded concurrency.** A document that pastes forty images should not
 *   open forty sockets - browsers queue past six per origin anyway, and Node
 *   will happily exhaust file descriptors. Four at a time, in document order.
 * - **Nothing fails the document.** A resolver that throws, rejects, returns a
 *   malformed `ResolvedImage`, or hangs past its own timeout costs exactly one
 *   picture.
 */

import { nodesOfType, type DocumentNode, type ImageNode, type ResolvedImage } from "../model.js";
import type { ImageMap, ImageResolver } from "../render/types.js";
import { isValidResolvedImage } from "./decode.js";
import { isObservableImageResolver } from "./resolver.js";
import { describeError } from "./transport.js";
import {
  createDiagnosticSink,
  diagnosticSeverity,
  type ImageDiagnostic,
  type ImageDiagnosticHandler,
} from "./types.js";

/** How many images to resolve at once. Matches a browser's per-origin socket budget. */
const DEFAULT_CONCURRENCY = 4;

/** Options for {@link resolveDocumentImages}. */
export interface ResolveDocumentImagesOptions {
  /** Maximum simultaneous `resolve()` calls. Defaults to 4. */
  readonly concurrency?: number | undefined;
  /** Called as each problem is found, in addition to being collected in the result. */
  readonly onDiagnostic?: ImageDiagnosticHandler | undefined;
}

/** What {@link resolveDocumentImages} produces. */
export interface DocumentImageResolution {
  /** Ready for `renderDocument(doc, { images })`. */
  readonly images: ImageMap;
  /**
   * Every problem found, in completion order - including the resolver's own,
   * when it is an
   * {@link import("./resolver.js").ObservableImageResolver} (everything
   * `createImageResolver` builds is).
   */
  readonly diagnostics: readonly ImageDiagnostic[];
  /** Distinct sources that produced bytes. */
  readonly resolvedCount: number;
  /** Distinct sources that did not. */
  readonly failedCount: number;
}

/**
 * Resolves every unresolved image in a document.
 *
 * Nodes that already carry `resolved` bytes (a plugin may have filled them in)
 * are skipped entirely - the resolver is never asked about them.
 *
 * @param doc - The document to walk.
 * @param resolver - Any {@link ImageResolver}; see `createImageResolver`.
 * @param options - Concurrency and diagnostics.
 * @returns The map to pass to the renderer, plus what went wrong. Never rejects.
 */
export async function resolveDocumentImages(
  doc: DocumentNode,
  resolver: ImageResolver,
  options: ResolveDocumentImagesOptions = {},
): Promise<DocumentImageResolution> {
  const diagnostics: ImageDiagnostic[] = [];
  const handler = options.onDiagnostic;
  const record = (diagnostic: ImageDiagnostic): void => {
    diagnostics.push(diagnostic);
    handler?.(diagnostic);
  };
  const sink = createDiagnosticSink(record);

  // Adopt the resolver's own diagnostics, if it has any to give, so the caller
  // has one place to look instead of two. Detached again in the `finally`
  // below: the resolver frequently outlives this call.
  const unobserve = isObservableImageResolver(resolver) ? resolver.observe(record) : null;

  const pending = nodesOfType(doc, "image").filter((node) => node.resolved === null);
  const bySource = new Map<string, ImageNode[]>();
  for (const node of pending) {
    const group = bySource.get(node.src);
    if (group === undefined) bySource.set(node.src, [node]);
    else group.push(node);
  }

  const sources = [...bySource.keys()];
  const images = new Map<ImageNode, ResolvedImage>();
  let resolvedCount = 0;
  let failedCount = 0;
  let next = 0;

  async function worker(): Promise<void> {
    for (;;) {
      const index = next;
      next += 1;
      if (index >= sources.length) return;
      const src = sources[index] as string;

      let resolved: ResolvedImage | null = null;
      try {
        resolved = await resolver.resolve(src);
      } catch (error: unknown) {
        // The contract says resolvers return null rather than throwing. Not
        // every implementation will, and one that does must not take the
        // document with it.
        sink.report("network-error", src, `image "${src}" resolver threw: ${describeError(error)}`);
        resolved = null;
      }

      if (resolved !== null && !isValidResolvedImage(resolved)) {
        sink.report(
          "invalid-resolution",
          src,
          `image "${src}": the resolver returned a ResolvedImage that OOXML cannot hold ` +
            `(an SVG needs a raster fallback, a raster must have none, and both need positive ` +
            `dimensions and non-empty bytes); it was discarded`,
        );
        resolved = null;
      }

      if (resolved === null) {
        failedCount += 1;
        continue;
      }

      resolvedCount += 1;
      for (const node of bySource.get(src) ?? []) images.set(node, resolved);
    }
  }

  const requested = options.concurrency;
  const concurrency = Math.max(
    1,
    Math.min(
      sources.length || 1,
      requested !== undefined && Number.isFinite(requested) && requested > 0
        ? Math.floor(requested)
        : DEFAULT_CONCURRENCY,
    ),
  );

  try {
    await Promise.all(Array.from({ length: concurrency }, () => worker()));
  } finally {
    unobserve?.();
  }

  return { images, diagnostics, resolvedCount, failedCount };
}

/**
 * Whether a resolution had any hard failure.
 *
 * Notices (an ignored EXIF orientation) do not count: the image still rendered.
 */
export function hasImageErrors(resolution: DocumentImageResolution): boolean {
  return resolution.diagnostics.some(
    (diagnostic) => diagnosticSeverity(diagnostic.code) === "error",
  );
}
