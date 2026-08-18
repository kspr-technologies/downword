/**
 * The async half of the pipeline.
 *
 * `renderDocument()` is synchronous by design, but the two things it needs from
 * the outside world — highlighted code and image bytes — are inherently async.
 * These two helpers do that work up front and hand back identity-keyed maps the
 * renderer can read without awaiting.
 *
 * Both are deliberately forgiving: an adapter that throws on one node loses
 * that node's enhancement (the code block renders plain, the image renders a
 * placeholder) and nothing else. A whole document must never fail to convert
 * because one syntax highlighter choked on one fence.
 *
 * **Forgiving is not the same as silent.** Every degradation is reported
 * through {@link PrepareOptions.onWarning} as an ordinary
 * {@link import("./types.js").RenderWarning}, on the same channel and in the
 * same vocabulary as the renderer's own — so an adapter that throws on *every*
 * node produces an unenhanced document *and* a stream of warnings saying why,
 * rather than a mystery.
 */

import {
  nodesOfType,
  type CodeBlockNode,
  type DocumentNode,
  type ImageNode,
  type ResolvedImage,
} from "../model.js";
import { reportRenderWarning } from "./types.js";
import type {
  Highlighter,
  HighlightMap,
  HighlightSpan,
  ImageMap,
  ImageResolver,
  PrepareOptions,
} from "./types.js";

/** Renders an unknown thrown value as one line, without assuming it is an Error. */
function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * Runs a {@link Highlighter} over every code block in the document.
 *
 * Blocks are processed in document order rather than in parallel: highlighters
 * are frequently stateful (highlight.js keeps a language registry, Shiki keeps
 * a loaded-grammar set) and ordered access keeps them reproducible.
 *
 * @param doc - The document to walk.
 * @param highlighter - Any {@link Highlighter}; see `createHighlighter`.
 * @param options - Where to send `highlighter-failed` warnings. See {@link PrepareOptions}.
 * @returns The map to pass to `renderDocument(doc, { highlights })`. Never rejects.
 */
export async function prepareHighlights(
  doc: DocumentNode,
  highlighter: Highlighter,
  options: PrepareOptions = {},
): Promise<HighlightMap> {
  const map = new Map<CodeBlockNode, readonly HighlightSpan[]>();

  for (const block of nodesOfType(doc, "codeBlock")) {
    if (map.has(block)) continue;
    try {
      map.set(block, await highlighter.highlight(block.value, block.lang));
    } catch (error: unknown) {
      // Leave the block unhighlighted; renderDocument falls back to plain
      // monospace, which is a correct - if duller - rendering.
      reportRenderWarning(
        options.onWarning,
        "highlighter-failed",
        `the highlighter threw on a ${block.lang ?? "plain"} code block; it renders unhighlighted: ${describe(error)}`,
      );
    }
  }

  return map;
}

/**
 * Runs an {@link ImageResolver} over every image in the document.
 *
 * Results are memoised by `src`, so the same asset referenced ten times is
 * fetched once. (docx then deduplicates the bytes again when packing, keying
 * media parts by content hash.)
 *
 * For a document-scale pass with bounded concurrency, resolver-contract
 * validation and the resolver's own diagnostics folded in, use
 * `resolveDocumentImages` from `@ksprtech/downword/images` instead; this is the minimal
 * version the renderer's own contract needs.
 *
 * @param doc - The document to walk.
 * @param resolver - Any {@link ImageResolver}; see `createImageResolver`.
 * @param options - Where to send `image-resolver-failed` warnings.
 * @returns The map to pass to `renderDocument(doc, { images })`. Never rejects.
 */
export async function prepareImages(
  doc: DocumentNode,
  resolver: ImageResolver,
  options: PrepareOptions = {},
): Promise<ImageMap> {
  const bySource = new Map<string, ResolvedImage | null>();
  const map = new Map<ImageNode, ResolvedImage>();

  for (const image of nodesOfType(doc, "image")) {
    if (image.resolved !== null) continue;

    let resolved = bySource.get(image.src);
    if (resolved === undefined) {
      try {
        resolved = await resolver.resolve(image.src);
      } catch (error: unknown) {
        // The contract says resolvers return null rather than throwing; one
        // that does costs exactly this picture, and says so.
        reportRenderWarning(
          options.onWarning,
          "image-resolver-failed",
          `image "${image.src}": the resolver threw, so it has no bytes: ${describe(error)}`,
        );
        resolved = null;
      }
      bySource.set(image.src, resolved);
    }

    if (resolved !== null) map.set(image, resolved);
  }

  return map;
}
