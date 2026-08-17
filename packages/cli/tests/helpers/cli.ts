/**
 * Spawning the built binary, and looking inside what it produced.
 *
 * stdout is captured as a **Buffer**, not a string: the CLI's whole stdin ->
 * stdout mode writes a zip, and decoding that as UTF-8 would corrupt it before
 * a single assertion ran.
 */

import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

/** `packages/cli`. */
export const PKG_DIR = fileURLToPath(new URL("../..", import.meta.url));

/** The binary under test. Built by `tests/global-setup.ts`. */
export const BIN = join(PKG_DIR, "dist", "index.js");

/** Committed markdown fixtures. */
export const FIXTURES = join(PKG_DIR, "tests", "__fixtures__", "markdown");

/**
 * Where generated `.docx` files go.
 *
 * `scripts/docx-validity.mjs` converts everything in this directory with
 * headless LibreOffice in CI, so anything written here is checked for
 * real-world openability and not merely for XML shape.
 */
export const OUT_DIR = join(PKG_DIR, "tests", "__fixtures__", "out");

/** Creates {@link OUT_DIR} and returns a path inside it. */
export function outPath(name: string): string {
  mkdirSync(OUT_DIR, { recursive: true });
  return join(OUT_DIR, name);
}

/** What a run of the CLI produced. */
export interface CliRun {
  readonly status: number | null;
  /** Raw bytes, because stdout may be a `.docx`. */
  readonly stdout: Buffer;
  /** stdout decoded as UTF-8, for the runs that print text. */
  readonly out: string;
  readonly stderr: string;
}

/** How to run it. */
export interface CliRunOptions {
  /** Fed to stdin. Omit for an empty, immediately-closed stdin. */
  readonly input?: string | Buffer;
  /** Working directory. Defaults to `packages/cli`. */
  readonly cwd?: string;
}

/** Runs the built binary to completion. */
export function runCli(args: readonly string[], options: CliRunOptions = {}): CliRun {
  const result = spawnSync(process.execPath, [BIN, ...args], {
    cwd: options.cwd ?? PKG_DIR,
    input: options.input ?? "",
    // 32 MiB: a .docx on stdout must never be silently truncated by the buffer.
    maxBuffer: 32 * 1024 * 1024,
  });
  const stdout = result.stdout ?? Buffer.alloc(0);
  return {
    status: result.status,
    stdout,
    out: stdout.toString("utf8"),
    stderr: (result.stderr ?? Buffer.alloc(0)).toString("utf8"),
  };
}

/** A `.docx`, opened. */
export interface Docx {
  /** Every entry in the zip, sorted. */
  readonly parts: readonly string[];
  /** One part as text. Throws when it is not there. */
  part(name: string): string;
  /** Whether a part exists. */
  has(name: string): boolean;
  /** `word/document.xml`. */
  readonly document: string;
  /** `word/document.xml` with the tags stripped: what a reader would see. */
  readonly text: string;
}

/** Opens a `.docx` from bytes. */
export async function openDocx(bytes: Uint8Array | Buffer): Promise<Docx> {
  const zip = await JSZip.loadAsync(bytes);
  const parts = Object.keys(zip.files).sort();
  const contents = new Map<string, string>();
  await Promise.all(
    parts.map(async (name) => {
      const file = zip.file(name);
      if (file !== null && !file.dir) contents.set(name, await file.async("string"));
    }),
  );

  const part = (name: string): string => {
    const found = contents.get(name);
    if (found === undefined) throw new Error(`no ${name} in the docx (have: ${parts.join(", ")})`);
    return found;
  };
  const document = part("word/document.xml");
  return {
    parts,
    part,
    has: (name: string) => contents.has(name),
    document,
    // Runs are separate elements, so this concatenates them; it is a
    // "does the word appear" check, not a layout check.
    text: document.replace(/<[^>]+>/g, ""),
  };
}

/** Opens a `.docx` from disk. */
export async function readDocx(path: string): Promise<Docx> {
  return openDocx(readFileSync(path));
}
