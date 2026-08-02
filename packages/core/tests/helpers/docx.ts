import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

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
