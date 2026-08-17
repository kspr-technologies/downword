import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { Packer, type Document } from "docx";
import JSZip from "jszip";

import type { DocumentNode } from "../../src/model.js";
import { renderDocument, type RenderOptions, type RenderWarning } from "../../src/render/index.js";

/**
 * Directory that golden tests write generated `.docx` files to.
 *
 * CI's LibreOffice gate (`scripts/docx-validity.mjs`) converts every file found
 * here to PDF, so anything written here is also checked for real-world
 * openability - not just for XML shape.
 */
export const FIXTURE_OUT_DIR = fileURLToPath(new URL("../__fixtures__/out/", import.meta.url));

/** Writes bytes into the fixture output directory and returns the absolute path. */
export async function writeFixture(name: string, bytes: Uint8Array): Promise<string> {
  await mkdir(FIXTURE_OUT_DIR, { recursive: true });
  const path = `${FIXTURE_OUT_DIR}${name}`;
  await writeFile(path, bytes);
  return path;
}

/** Reads a single part out of a `.docx` zip as UTF-8 text. */
export async function readDocxPart(bytes: Uint8Array, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file(part);
  if (file === null) {
    throw new Error(`docx part not found: ${part} (have: ${Object.keys(zip.files).join(", ")})`);
  }
  return file.async("string");
}

/** Sorted list of every entry in a `.docx` zip. */
export async function listDocxParts(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  return Object.keys(zip.files).sort();
}

/** Packs a docx `Document` into `.docx` bytes. */
export async function packDocument(file: Document): Promise<Uint8Array> {
  return new Uint8Array(await Packer.toArrayBuffer(file));
}

/** The OOXML parts the renderer's tests care about. */
export interface RenderedParts {
  readonly bytes: Uint8Array;
  readonly document: string;
  readonly styles: string;
  readonly numbering: string | null;
  readonly footnotes: string | null;
  /** `word/footer1.xml`; absent unless the section declares a footer. */
  readonly footer: string | null;
  /** `word/settings.xml`, which is where `<w:updateFields/>` lands. */
  readonly settings: string | null;
  /** `docProps/core.xml`: title, author, keywords, description. */
  readonly coreProperties: string | null;
  /** `docProps/custom.xml`: everything frontmatter had no core property for. */
  readonly customProperties: string | null;
  readonly warnings: readonly RenderWarning[];
}

/**
 * Renders a model document, packs it and unzips the interesting parts.
 *
 * Collecting warnings here (rather than in each test) keeps the renderer's
 * `onWarning` contract exercised on every single golden case, so a rendering
 * that silently starts dropping content shows up as a warning diff.
 */
export async function renderParts(
  doc: DocumentNode,
  options: RenderOptions = {},
): Promise<RenderedParts> {
  const warnings: RenderWarning[] = [];
  const file = renderDocument(doc, {
    ...options,
    onWarning: (warning) => {
      warnings.push(warning);
      options.onWarning?.(warning);
    },
  });

  const bytes = await packDocument(file);
  const zip = await JSZip.loadAsync(bytes);
  const optional = async (part: string): Promise<string | null> => {
    const entry = zip.file(part);
    return entry === null ? null : entry.async("string");
  };

  return {
    bytes,
    document: await readDocxPart(bytes, "word/document.xml"),
    styles: await readDocxPart(bytes, "word/styles.xml"),
    numbering: await optional("word/numbering.xml"),
    footnotes: await optional("word/footnotes.xml"),
    footer: await optional("word/footer1.xml"),
    settings: await optional("word/settings.xml"),
    coreProperties: await optional("docProps/core.xml"),
    customProperties: await optional("docProps/custom.xml"),
    warnings,
  };
}
