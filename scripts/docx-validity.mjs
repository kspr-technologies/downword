#!/usr/bin/env node
/**
 * T0.4 - .docx validity gate.
 *
 * Every `.docx` this repo generates must be openable by a real Office
 * implementation, not merely well-shaped XML. This script:
 *
 *   1. Structurally checks every generated `.docx` (zip magic + the two
 *      mandatory OOXML parts). Runs everywhere, needs no LibreOffice.
 *   2. Converts every generated `.docx` to PDF with headless LibreOffice and
 *      fails if any conversion errors or produces no readable PDF.
 *   3. Runs a NEGATIVE CONTROL against a deliberately corrupt fixture
 *      (packages/core/tests/__fixtures__/corrupt/corrupt.docx). If that file
 *      passes, the gate itself is broken and the run fails - a gate that never
 *      says "no" is worse than no gate.
 *
 * Usage:
 *   node scripts/docx-validity.mjs [--require-soffice] [--keep]
 *
 * Behaviour when LibreOffice is missing:
 *   - locally: steps 1 and 3a still run, step 2 is reported as SKIPPED, exit 0.
 *   - in CI (`$CI` set) or with --require-soffice: hard failure, exit 1.
 *
 * Env:
 *   SOFFICE   absolute path to the `soffice` binary (overrides discovery)
 */

import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  openSync,
  readSync,
  closeSync,
  readdirSync,
  rmSync,
  statSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

/** Directories that generated .docx fixtures land in. */
const GENERATED_DIRS = [
  "packages/core/tests/__fixtures__/out",
  "packages/cli/tests/__fixtures__/out",
  "examples/basic-node/out",
  "examples/basic-cjs/out",
];

/** The deliberately-broken file used to prove the gate can fail. */
const CORRUPT_FIXTURE = "packages/core/tests/__fixtures__/corrupt/corrupt.docx";

/** Candidate locations for the LibreOffice binary. */
const SOFFICE_CANDIDATES = [
  "/usr/bin/soffice",
  "/usr/local/bin/soffice",
  "/opt/homebrew/bin/soffice",
  "/snap/bin/libreoffice",
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
];

const args = new Set(process.argv.slice(2));
const REQUIRE_SOFFICE = args.has("--require-soffice") || Boolean(process.env["CI"]);
const KEEP_OUTPUT = args.has("--keep");

const results = [];
let failed = false;

function record(status, name, detail) {
  results.push({ status, name, detail });
  if (status === "FAIL") failed = true;
  const icon = { PASS: "  ok  ", FAIL: " FAIL ", SKIP: " skip " }[status];
  console.log(`[${icon}] ${name}${detail ? ` - ${detail}` : ""}`);
}

/** Reads the first `length` bytes of a file. */
function readMagic(file, length) {
  const fd = openSync(file, "r");
  try {
    const buffer = Buffer.alloc(length);
    const read = readSync(fd, buffer, 0, length, 0);
    return buffer.subarray(0, read);
  } finally {
    closeSync(fd);
  }
}

/**
 * Structural OOXML check without a zip library.
 *
 * A `.docx` is a zip; zip local file headers store entry names uncompressed, so
 * the mandatory part names are always findable as raw bytes.
 */
function inspectDocx(file) {
  const problems = [];
  const size = statSync(file).size;

  if (size === 0) problems.push("file is empty");

  const magic = readMagic(file, 4);
  if (!(magic[0] === 0x50 && magic[1] === 0x4b && magic[2] === 0x03 && magic[3] === 0x04)) {
    problems.push(
      `not a zip (magic ${[...magic].map((b) => b.toString(16).padStart(2, "0")).join(" ")})`,
    );
  }

  const fd = openSync(file, "r");
  let haystack;
  try {
    const buffer = Buffer.alloc(size);
    readSync(fd, buffer, 0, size, 0);
    haystack = buffer.toString("latin1");
  } finally {
    closeSync(fd);
  }

  for (const part of ["[Content_Types].xml", "word/document.xml"]) {
    if (!haystack.includes(part)) problems.push(`missing OOXML part: ${part}`);
  }

  return { size, problems };
}

/** Finds the LibreOffice binary, or returns null. */
function findSoffice() {
  const fromEnv = process.env["SOFFICE"];
  if (fromEnv && existsSync(fromEnv)) return fromEnv;

  const which = spawnSync("which", ["soffice"], { encoding: "utf8" });
  if (which.status === 0) {
    const path = which.stdout.trim();
    if (path && existsSync(path)) return path;
  }

  return SOFFICE_CANDIDATES.find((candidate) => existsSync(candidate)) ?? null;
}

/**
 * Converts one file to PDF with headless LibreOffice.
 *
 * `-env:UserInstallation` gives every invocation a private profile; without it
 * concurrent or repeated runs silently reuse a locked profile and "succeed"
 * while producing nothing.
 *
 * @param {string} soffice
 * @param {string} file
 * @param {string} outDir
 * @param {string[]} extraArgs e.g. forcing an import filter
 */
function convertToPdf(soffice, file, outDir, extraArgs = []) {
  mkdirSync(outDir, { recursive: true });
  const profile = mkdtempSync(join(tmpdir(), "downword-lo-"));

  const argv = [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--headless",
    "--norestore",
    "--invisible",
    "--nolockcheck",
    "--nodefault",
    "--nofirststartwizard",
    ...extraArgs,
    "--convert-to",
    "pdf",
    "--outdir",
    outDir,
    file,
  ];

  const proc = spawnSync(soffice, argv, { encoding: "utf8", timeout: 180_000 });

  const base = file
    .split("/")
    .pop()
    .replace(/\.docx$/i, "");
  const pdf = join(outDir, `${base}.pdf`);

  let pdfOk = false;
  let pdfSize = 0;
  if (existsSync(pdf)) {
    pdfSize = statSync(pdf).size;
    pdfOk = pdfSize > 0 && readMagic(pdf, 4).toString("latin1") === "%PDF";
  }

  rmSync(profile, { recursive: true, force: true });

  return {
    ok: proc.status === 0 && pdfOk,
    exitCode: proc.status,
    pdf,
    pdfSize,
    stderr: (proc.stderr ?? "").trim(),
    stdout: (proc.stdout ?? "").trim(),
  };
}

/** Collects every generated .docx under the known output directories. */
function collectGenerated() {
  const found = [];
  for (const dir of GENERATED_DIRS) {
    const abs = join(REPO_ROOT, dir);
    if (!existsSync(abs)) continue;
    for (const entry of readdirSync(abs, { withFileTypes: true })) {
      if (entry.isFile() && entry.name.toLowerCase().endsWith(".docx")) {
        found.push(join(abs, entry.name));
      }
    }
  }
  return found.sort();
}

// ---------------------------------------------------------------------------

console.log("downword .docx validity gate\n");

const generated = collectGenerated();
const corruptPath = join(REPO_ROOT, CORRUPT_FIXTURE);

if (generated.length === 0) {
  console.error(
    "No generated .docx files found.\n" +
      "Run `pnpm test` (and optionally `pnpm build && pnpm examples`) first; they write into:\n" +
      GENERATED_DIRS.map((dir) => `  - ${dir}/`).join("\n"),
  );
  process.exit(1);
}

console.log(`Found ${generated.length} generated .docx file(s).\n`);

// --- 1. structural checks --------------------------------------------------
console.log("1) Structural checks (no LibreOffice required)");
for (const file of generated) {
  const rel = relative(REPO_ROOT, file);
  const { size, problems } = inspectDocx(file);
  if (problems.length === 0) {
    record("PASS", rel, `${size} bytes, valid OOXML container`);
  } else {
    record("FAIL", rel, problems.join("; "));
  }
}

// --- 1b. negative control, structural --------------------------------------
console.log("\n2) Negative control - the corrupt fixture must NOT look like a .docx");
if (!existsSync(corruptPath)) {
  record("FAIL", CORRUPT_FIXTURE, "fixture is missing; the gate cannot prove it can fail");
} else {
  const { problems } = inspectDocx(corruptPath);
  if (problems.length > 0) {
    record("PASS", CORRUPT_FIXTURE, `correctly rejected (${problems.join("; ")})`);
  } else {
    record(
      "FAIL",
      CORRUPT_FIXTURE,
      "structural check accepted a corrupt file - THE GATE IS BROKEN",
    );
  }
}

// --- 2. LibreOffice conversions --------------------------------------------
console.log("\n3) LibreOffice headless conversion");
const soffice = findSoffice();

if (soffice === null) {
  const message =
    "soffice not found. Install LibreOffice, or set SOFFICE=/path/to/soffice.\n" +
    "  macOS : brew install --cask libreoffice\n" +
    "  Debian: sudo apt-get install -y libreoffice-writer\n" +
    "  Looked in: " +
    [process.env["SOFFICE"] ?? "$SOFFICE (unset)", "PATH", ...SOFFICE_CANDIDATES].join(", ");

  if (REQUIRE_SOFFICE) {
    record("FAIL", "libreoffice", message);
  } else {
    record("SKIP", "libreoffice", message);
    console.log(
      "\n  Conversion checks were skipped locally. CI runs them with --require-soffice,\n" +
        "  so a document that LibreOffice cannot open will still be caught before merge.",
    );
  }
} else {
  console.log(`   using: ${soffice}\n`);
  const outRoot = mkdtempSync(join(tmpdir(), "downword-pdf-"));

  for (const file of generated) {
    const rel = relative(REPO_ROOT, file);
    const result = convertToPdf(soffice, file, join(outRoot, "valid"));
    if (result.ok) {
      record("PASS", rel, `-> pdf (${result.pdfSize} bytes)`);
    } else {
      record(
        "FAIL",
        rel,
        `soffice exit=${result.exitCode}, pdf=${existsSync(result.pdf) ? `${result.pdfSize} bytes` : "not produced"}` +
          (result.stderr ? `\n         stderr: ${result.stderr}` : "") +
          (result.stdout ? `\n         stdout: ${result.stdout}` : ""),
      );
    }
  }

  // Negative control through LibreOffice. The MS Word 2007 XML import filter is
  // forced: without it LibreOffice sniffs the content, falls back to its plain
  // text filter and happily "converts" a text file, which would make the
  // negative control meaningless.
  console.log("\n4) Negative control - LibreOffice must REFUSE the corrupt fixture");
  if (existsSync(corruptPath)) {
    const result = convertToPdf(soffice, corruptPath, join(outRoot, "corrupt"), [
      "--infilter=MS Word 2007 XML",
    ]);
    if (result.ok) {
      record(
        "FAIL",
        CORRUPT_FIXTURE,
        "LibreOffice produced a PDF from a corrupt .docx - THE GATE IS BROKEN",
      );
    } else {
      record("PASS", CORRUPT_FIXTURE, `correctly refused (exit=${result.exitCode}, no valid pdf)`);
    }
  }

  if (KEEP_OUTPUT) {
    console.log(`\n   PDFs kept in ${outRoot}`);
  } else {
    rmSync(outRoot, { recursive: true, force: true });
  }
}

// ---------------------------------------------------------------------------
const counts = results.reduce((acc, r) => ({ ...acc, [r.status]: (acc[r.status] ?? 0) + 1 }), {});
console.log(
  `\nSummary: ${counts["PASS"] ?? 0} passed, ${counts["FAIL"] ?? 0} failed, ${counts["SKIP"] ?? 0} skipped`,
);

process.exit(failed ? 1 : 0);
