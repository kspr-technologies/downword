/**
 * The pipeline: markdown in, `.docx` out.
 *
 * ```text
 *   parse ──▶ resolve images ─┐
 *                             ├──▶ render ──▶ pack
 *             highlight ──────┘
 * ```
 *
 * `parseMarkdown` and `renderDocument` are both synchronous and pure; the two
 * things that are not — fetching image bytes and loading a syntax grammar — run
 * in the middle, in parallel, and hand the renderer identity-keyed maps it can
 * read without awaiting. That split is why the same code runs unchanged in a
 * browser tab, a worker, Node and a server handler.
 *
 * Nothing here performs I/O of its own accord: the default image resolver reads
 * `data:` and `blob:` URLs only, and no highlighter means highlight.js is never
 * loaded. A conversion is an offline, in-memory operation unless you opt out.
 */

import { Packer, type Document } from "docx";

import { DownwordError, wrapError } from "./errors.js";
import { createImageResolver, resolveDocumentImages } from "./images/index.js";
import type { DocumentNode } from "./model.js";
import {
  resolveConvertOptions,
  type ConvertOptions,
  type ConvertWarning,
  type ConvertWarningHandler,
  type ResolvedConvertOptions,
} from "./options.js";
import { parseMarkdown } from "./parse/index.js";
import {
  prepareHighlights,
  renderDocument,
  type HighlightMap,
  type ImageMap,
} from "./render/index.js";

/** Version of this package, injected at build time. See `types/globals.d.ts`. */
export const VERSION: string = __DOWNWORD_VERSION__;

/**
 * MIME type of an OOXML WordprocessingML document.
 *
 * The value a `Blob`, a `Content-Type` header or an `<a download>` needs. Word
 * will open a `.docx` served as `application/octet-stream`, but browsers and
 * mail clients treat it as an unknown binary and say so.
 */
export const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** Warning handler that drops everything. */
const IGNORE_WARNINGS: ConvertWarningHandler = () => {};

/**
 * Wraps the caller's handler so that a logger which throws cannot fail a
 * conversion — losing a warning is always better than losing the document.
 */
function createEmitter(handler: ConvertWarningHandler | undefined): ConvertWarningHandler {
  if (handler === undefined) return IGNORE_WARNINGS;
  return (warning: ConvertWarning) => {
    try {
      handler(warning);
    } catch {
      /* a host whose logger throws is not this document's problem */
    }
  };
}

/** Runs the image pass, folding its diagnostics into the shared warning stream. */
async function resolveImages(
  document: DocumentNode,
  resolved: ResolvedConvertOptions,
  emit: ConvertWarningHandler,
): Promise<ImageMap> {
  const resolver =
    resolved.imageResolver ?? createImageResolver({ allowRemote: resolved.allowRemoteImages });

  const { images } = await resolveDocumentImages(document, resolver, {
    onDiagnostic: (diagnostic) => {
      emit({ ...diagnostic, stage: "image" });
    },
  });
  return images;
}

/**
 * Runs the highlighting pass, or returns an empty map when there is no highlighter.
 *
 * Its warnings join the same stream as the renderer's, under the same `render`
 * stage: a fence that lost its colour because the adapter threw is the same
 * kind of news to a host as one that lost it because the renderer could not use
 * the result.
 */
async function resolveHighlights(
  document: DocumentNode,
  resolved: ResolvedConvertOptions,
  emit: ConvertWarningHandler,
): Promise<HighlightMap> {
  if (resolved.highlighter === null) return new Map();
  return prepareHighlights(document, resolved.highlighter, {
    onWarning: (warning) => {
      emit({ ...warning, stage: "render" });
    },
  });
}

/**
 * Converts markdown into a `docx` `Document`, one step short of the zip.
 *
 * The escape hatch for anything {@link convert} does not do: pack it yourself
 * with `Packer.toBase64String` / `toStream`, or inspect what was produced. It
 * *is* `convert()` minus the final `Packer` call.
 *
 * @param markdown - The source. May be empty; the result is a valid, empty document.
 * @param options - See {@link ConvertOptions}.
 * @returns The `docx` `Document` (the library's `File` object).
 * @throws {DownwordError} with code `"invalid-input"`, `"invalid-options"`,
 *   `"parse-failed"` or `"render-failed"`. Never throws anything else.
 */
export async function convertToDocument(
  markdown: string,
  options: ConvertOptions = {},
): Promise<Document> {
  if (typeof markdown !== "string") {
    throw new DownwordError(
      "invalid-input",
      `markdown must be a string, got ${markdown === null ? "null" : typeof markdown}`,
    );
  }

  const emit = createEmitter(options?.onWarning);
  const resolved = resolveConvertOptions(options, emit);

  let document: DocumentNode;
  try {
    document = parseMarkdown(markdown, resolved.parse);
  } catch (cause: unknown) {
    // parseMarkdown itself never throws; a plugin from options.plugins can.
    throw wrapError("parse-failed", "could not parse the markdown", cause);
  }

  // Independent of each other, and both I/O-bound. A document with no images
  // and no highlighter settles both on the first microtask.
  const [images, highlights] = await Promise.all([
    resolveImages(document, resolved, emit),
    resolveHighlights(document, resolved, emit),
  ]);

  try {
    return renderDocument(document, { ...resolved.render, images, highlights });
  } catch (cause: unknown) {
    throw wrapError("render-failed", "could not render the document", cause);
  }
}

/**
 * Converts markdown to a Word `.docx` file.
 *
 * ```ts
 * import { convert } from "@ksprtech/downword";
 *
 * const bytes = await convert("# Hello\n\nFrom **downword**.");
 * ```
 *
 * Runs identically in the browser and in Node, and performs no I/O: a network
 * request is impossible unless {@link ConvertOptions.allowRemoteImages} is set
 * or you supply your own {@link ConvertOptions.imageResolver}.
 *
 * Malformed markdown is not an error — markdown has no syntax errors, only
 * constructs OOXML cannot hold, and those are reported through
 * {@link ConvertOptions.onWarning} while the conversion continues. Converting
 * `""` produces a valid, empty document rather than throwing.
 *
 * @param markdown - The source. CommonMark plus GFM tables, strikethrough,
 *   task lists and autolinks, plus footnotes.
 * @param options - See {@link ConvertOptions}. Every field has a default.
 * @returns The `.docx` bytes. Write them to disk, or wrap them in a `Blob`
 *   tagged {@link DOCX_MIME_TYPE} — or call {@link convertToBlob}, which does
 *   that for you.
 * @throws {DownwordError} and nothing else. See `DownwordErrorCode`.
 */
export async function convert(markdown: string, options: ConvertOptions = {}): Promise<Uint8Array> {
  const file = await convertToDocument(markdown, options);
  try {
    // toArrayBuffer is the one Packer output that is native in both runtimes.
    // toBuffer and toStream also "work" in a browser, but only by way of the
    // Buffer and readable-stream polyfills docx bundles, which is not something
    // to hand a caller.
    return new Uint8Array(await Packer.toArrayBuffer(file));
  } catch (cause: unknown) {
    throw wrapError("pack-failed", "could not build the .docx package", cause);
  }
}

/**
 * Converts markdown to a `Blob` tagged with {@link DOCX_MIME_TYPE}.
 *
 * The browser convenience form — a `Blob` is what `URL.createObjectURL`,
 * `showSaveFilePicker` and `fetch(url, { body })` all want:
 *
 * ```ts
 * import { convertToBlob } from "@ksprtech/downword";
 *
 * const blob = await convertToBlob("# Hello");
 * const url = URL.createObjectURL(blob);
 * ```
 *
 * Also works in Node ≥18, where `Blob` is a global.
 *
 * @param markdown - The source.
 * @param options - See {@link ConvertOptions}.
 * @returns A `Blob` of the `.docx`, with the correct `type`.
 * @throws {DownwordError} and nothing else.
 */
export async function convertToBlob(markdown: string, options: ConvertOptions = {}): Promise<Blob> {
  const file = await convertToDocument(markdown, options);
  try {
    return await Packer.toBlob(file);
  } catch (cause: unknown) {
    throw wrapError("pack-failed", "could not build the .docx package", cause);
  }
}
