#!/usr/bin/env node
/**
 * Packaging smoke test.
 *
 * `pnpm test` only ever loads TypeScript source. This script checks the thing
 * users actually get: it packs `packages/core` into a tarball, installs that
 * tarball into a throwaway project outside the workspace, and imports it both
 * as ESM and as CJS, asserting the export shape of the main entry and of every
 * subpath export.
 *
 * Catches: a broken `exports` map, missing `dist` files, a `files` field that
 * forgets something, `.d.ts` paths that do not exist, and ESM/CJS interop bugs.
 *
 * Usage:
 *   node scripts/pack-smoke.mjs [--dir <scratch dir>] [--keep]
 */

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const CORE_DIR = join(REPO_ROOT, "packages", "core");

const argv = process.argv.slice(2);
const dirFlag = argv.indexOf("--dir");
const KEEP = argv.includes("--keep");
const SCRATCH =
  dirFlag !== -1 && argv[dirFlag + 1]
    ? resolve(argv[dirFlag + 1])
    : join(tmpdir(), `downword-pack-smoke-${Date.now()}`);

function run(command, args, options = {}) {
  console.log(`$ ${command} ${args.join(" ")}`);
  return execFileSync(command, args, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "inherit"],
    ...options,
  });
}

// 1. Build + pack -----------------------------------------------------------
console.log("\n== 1. build and pack packages/core ==");
run("pnpm", ["--filter", "downword", "build"], { cwd: REPO_ROOT });

rmSync(SCRATCH, { recursive: true, force: true });
mkdirSync(SCRATCH, { recursive: true });

run("pnpm", ["pack", "--pack-destination", SCRATCH], { cwd: CORE_DIR });

const tarball = readdirSync(SCRATCH).find((file) => file.endsWith(".tgz"));
if (!tarball) throw new Error(`no tarball produced in ${SCRATCH}`);
const tarballPath = join(SCRATCH, tarball);
console.log(`   tarball: ${tarballPath}`);

// 2. Install the tarball into a throwaway project ---------------------------
console.log("\n== 2. install the tarball into a scratch project ==");
const project = join(SCRATCH, "project");
mkdirSync(project, { recursive: true });

writeFileSync(
  join(project, "package.json"),
  `${JSON.stringify(
    {
      name: "downword-pack-smoke",
      version: "0.0.0",
      private: true,
      type: "module",
    },
    null,
    2,
  )}\n`,
);

run("npm", ["install", "--no-audit", "--no-fund", "--loglevel", "error", tarballPath], {
  cwd: project,
});

// 3. Consumer scripts -------------------------------------------------------
console.log("\n== 3. import as ESM and require as CJS ==");

const ESM_CHECK = `
import assert from "node:assert/strict";
import * as downword from "downword";
import { mathPlugin } from "downword/plugins/math";
import { mermaidPlugin } from "downword/plugins/mermaid";

assert.equal(typeof downword.convert, "function", "convert must be a function");
assert.equal(typeof downword.VERSION, "string", "VERSION must be a string");
assert.match(downword.VERSION, /^\\d+\\.\\d+\\.\\d+/, "VERSION must look like semver");
assert.equal(
  downword.DOCX_MIME_TYPE,
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
);
assert.deepEqual(
  Object.keys(downword).sort(),
  ["DOCX_MIME_TYPE", "VERSION", "convert"],
  "unexpected ESM export shape",
);
assert.equal(typeof mathPlugin, "function", "downword/plugins/math must export mathPlugin");
assert.equal(typeof mermaidPlugin, "function", "downword/plugins/mermaid must export mermaidPlugin");

const result = await downword.convert("# Hello\\n\\nfrom the packed tarball.");
assert.ok(result.bytes instanceof Uint8Array, "bytes must be a Uint8Array");
assert.deepEqual([...result.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "must be a zip");
assert.ok(result.bytes.byteLength > 1000, "docx looks too small");
assert.ok(Array.isArray(result.warnings), "warnings must be an array");

console.log("[esm] ok - downword@" + downword.VERSION + ", " + result.bytes.byteLength + " bytes");
`;

const CJS_CHECK = `
"use strict";
const assert = require("node:assert/strict");
const downword = require("downword");
const math = require("downword/plugins/math");
const mermaid = require("downword/plugins/mermaid");

assert.equal(typeof downword.convert, "function", "convert must be a function");
assert.equal(typeof downword.VERSION, "string", "VERSION must be a string");
assert.equal(
  downword.DOCX_MIME_TYPE,
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
);
assert.deepEqual(
  Object.keys(downword).sort(),
  ["DOCX_MIME_TYPE", "VERSION", "convert"],
  "unexpected CJS export shape",
);
assert.equal(typeof math.mathPlugin, "function", "downword/plugins/math must export mathPlugin");
assert.equal(
  typeof mermaid.mermaidPlugin,
  "function",
  "downword/plugins/mermaid must export mermaidPlugin",
);

downword
  .convert("# Hello\\n\\nfrom require().")
  .then((result) => {
    assert.ok(result.bytes instanceof Uint8Array, "bytes must be a Uint8Array");
    assert.deepEqual([...result.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "must be a zip");
    console.log("[cjs] ok - downword@" + downword.VERSION + ", " + result.bytes.byteLength + " bytes");
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
`;

const TYPES_CHECK = `
const assert = require("node:assert/strict");
const { existsSync } = require("node:fs");
const { createRequire } = require("node:module");
const { dirname, join } = require("node:path");

const require_ = createRequire(join(process.cwd(), "index.js"));
const pkgPath = require_.resolve("downword/package.json");
const pkg = require_("downword/package.json");
const root = dirname(pkgPath);

for (const [entry, conditions] of Object.entries(pkg.exports)) {
  if (typeof conditions === "string") continue;
  for (const [condition, target] of Object.entries(conditions)) {
    const files = typeof target === "string" ? { default: target } : target;
    for (const [inner, file] of Object.entries(files)) {
      const abs = join(root, file);
      assert.ok(existsSync(abs), \`missing \${entry} -> \${condition}.\${inner}: \${file}\`);
    }
  }
}
console.log("[files] ok - every path in the exports map exists in the tarball");
`;

writeFileSync(join(project, "esm-check.mjs"), ESM_CHECK);
writeFileSync(join(project, "cjs-check.cjs"), CJS_CHECK);
writeFileSync(join(project, "files-check.cjs"), TYPES_CHECK);

for (const script of ["esm-check.mjs", "cjs-check.cjs", "files-check.cjs"]) {
  execFileSync(process.execPath, [script], { cwd: project, stdio: "inherit" });
}

// 4. Done -------------------------------------------------------------------
if (KEEP || !existsSync(SCRATCH)) {
  console.log(`\npack smoke PASSED. Scratch project kept at ${SCRATCH}`);
} else {
  rmSync(SCRATCH, { recursive: true, force: true });
  console.log("\npack smoke PASSED.");
}
