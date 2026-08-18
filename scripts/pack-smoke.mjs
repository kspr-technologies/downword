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
run("pnpm", ["--filter", "@ksprtech/downword", "build"], { cwd: REPO_ROOT });

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

/**
 * Runtime exports the published main entry must always carry.
 *
 * Deliberately a *subset* check rather than a deepEqual of every key: adding an
 * export is not a breaking change, and pinning the full list here would turn
 * every additive release into a red CI run. Removing one of these, however, is
 * breaking - so these are the ones worth nailing down.
 */
const REQUIRED_EXPORTS = [
  "DOCX_MIME_TYPE",
  "DEFAULT_THEME",
  "DownwordError",
  "NULL_IMAGE_RESOLVER",
  "THEMES",
  "VERSION",
  "convert",
  "convertToBlob",
  "convertToDocument",
  "createImageResolver",
  "isDownwordError",
  "parseMarkdown",
  "renderDocument",
];

const ESM_CHECK = `
import assert from "node:assert/strict";
import * as downword from "@ksprtech/downword";
import { createHighlighter } from "@ksprtech/downword/highlight";
import { createNodeImageResolver } from "@ksprtech/downword/images/node";
import { convertWithMath, mathPlugin } from "@ksprtech/downword/plugins/math";
import { renderMermaid, countMermaidDiagrams } from "@ksprtech/downword/plugins/mermaid";

const required = ${JSON.stringify(REQUIRED_EXPORTS)};
const missing = required.filter((name) => !(name in downword));
assert.deepEqual(missing, [], "downword is missing required exports");

assert.equal(typeof downword.convert, "function", "convert must be a function");
assert.match(downword.VERSION, /^\\d+\\.\\d+\\.\\d+/, "VERSION must look like semver");
assert.equal(
  downword.DOCX_MIME_TYPE,
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
);
assert.equal(typeof createHighlighter, "function", "@ksprtech/downword/highlight must export createHighlighter");
assert.equal(
  typeof createNodeImageResolver,
  "function",
  "@ksprtech/downword/images/node must export createNodeImageResolver",
);
assert.equal(typeof mathPlugin, "function", "@ksprtech/downword/plugins/math must export mathPlugin");
assert.equal(typeof renderMermaid, "function", "@ksprtech/downword/plugins/mermaid must export renderMermaid");
assert.equal(
  countMermaidDiagrams({ type: "document", metadata: {}, children: [] }),
  0,
  "countMermaidDiagrams must work without a DOM",
);

const warnings = [];
const bytes = await downword.convert("# Hello\\n\\nfrom the packed tarball.\\n\\n- a\\n- b", {
  metadata: { title: "pack smoke" },
  onWarning: (warning) => warnings.push(warning),
});
assert.ok(bytes instanceof Uint8Array, "convert must resolve to a Uint8Array");
assert.deepEqual([...bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "must be a zip");
assert.ok(bytes.byteLength > 1000, "docx looks too small");
assert.deepEqual(warnings, [], "clean markdown must not warn");

const blob = await downword.convertToBlob("# Hello");
assert.equal(blob.type, downword.DOCX_MIME_TYPE, "convertToBlob must tag the mime type");

// Both plugin entries are *driven*, not merely imported. This scratch project
// deliberately has neither temml/mathml2omml nor a DOM, so it pins the
// documented degradation: an equation becomes its TeX source and a diagram
// stays a code block, and in both cases a real .docx still comes out.
const mathWarnings = [];
const mathBytes = await convertWithMath("# Math\\n\\nInline $E=mc^2$ and:\\n\\n$$\\\\frac{1}{3}$$\\n", {
  math: { onWarning: (warning) => mathWarnings.push(warning.code) },
});
assert.deepEqual([...mathBytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "math docx must be a zip");
assert.ok(mathBytes.byteLength > 1000, "math docx looks too small");
assert.ok(
  mathWarnings.includes("engine-unavailable"),
  "without temml/mathml2omml the math plugin must say so, not fail silently: " +
    JSON.stringify(mathWarnings),
);

const parsed = downword.parseMarkdown("# Diagram\\n\\n\`\`\`mermaid\\nflowchart LR\\n  A --> B\\n\`\`\`\\n");
assert.equal(countMermaidDiagrams(parsed), 1, "the fence must be seen as a diagram");
const rendered = await renderMermaid(parsed);
assert.equal(rendered.rendered, 0, "no DOM here, so nothing may render");
assert.deepEqual(
  rendered.warnings.map((warning) => warning.code),
  ["no-dom"],
  "a DOM-less runtime must report exactly no-dom",
);
const mermaidBytes = await downword.convert(
  "# Diagram\\n\\n\`\`\`mermaid\\nflowchart LR\\n  A --> B\\n\`\`\`\\n",
);
assert.deepEqual([...mermaidBytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "mermaid docx is a zip");

console.log(
  "[esm] ok - downword@" + downword.VERSION + ", " + bytes.byteLength + " bytes; " +
    "math " + mathBytes.byteLength + " bytes, mermaid " + mermaidBytes.byteLength + " bytes",
);
`;

const CJS_CHECK = `
"use strict";
const assert = require("node:assert/strict");
const downword = require("@ksprtech/downword");
const highlight = require("@ksprtech/downword/highlight");
const nodeImages = require("@ksprtech/downword/images/node");
const math = require("@ksprtech/downword/plugins/math");
const mermaid = require("@ksprtech/downword/plugins/mermaid");

const required = ${JSON.stringify(REQUIRED_EXPORTS)};
const missing = required.filter((name) => !(name in downword));
assert.deepEqual(missing, [], "downword is missing required exports");

assert.equal(typeof downword.convert, "function", "convert must be a function");
assert.equal(typeof downword.VERSION, "string", "VERSION must be a string");
assert.equal(
  downword.DOCX_MIME_TYPE,
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
);
assert.equal(
  typeof highlight.createHighlighter,
  "function",
  "@ksprtech/downword/highlight must export createHighlighter",
);
assert.equal(
  typeof nodeImages.createNodeImageResolver,
  "function",
  "@ksprtech/downword/images/node must export createNodeImageResolver",
);
assert.equal(typeof math.mathPlugin, "function", "@ksprtech/downword/plugins/math must export mathPlugin");
assert.equal(
  typeof mermaid.renderMermaid,
  "function",
  "@ksprtech/downword/plugins/mermaid must export renderMermaid",
);

async function main() {
  const bytes = await downword.convert("# Hello\\n\\nfrom require().");
  assert.ok(bytes instanceof Uint8Array, "convert must resolve to a Uint8Array");
  assert.deepEqual([...bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "must be a zip");

  // The same two plugin entries, driven through require(). A subpath that
  // resolves under import but not under require is a real and common packaging
  // bug, and importing without calling would not catch a broken interop shape.
  const mathBytes = await math.convertWithMath("# Math\\n\\n$E=mc^2$\\n");
  assert.deepEqual([...mathBytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04], "math docx must be a zip");

  const parsed = downword.parseMarkdown(
    "# Diagram\\n\\n\`\`\`mermaid\\nflowchart LR\\n  A --> B\\n\`\`\`\\n",
  );
  assert.equal(mermaid.countMermaidDiagrams(parsed), 1, "the fence must be seen as a diagram");
  const rendered = await mermaid.renderMermaid(parsed);
  assert.deepEqual(
    rendered.warnings.map((warning) => warning.code),
    ["no-dom"],
    "a DOM-less runtime must report exactly no-dom",
  );

  console.log(
    "[cjs] ok - downword@" + downword.VERSION + ", " + bytes.byteLength + " bytes; " +
      "math " + mathBytes.byteLength + " bytes",
  );
}

main().catch((error) => {
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
const pkgPath = require_.resolve("@ksprtech/downword/package.json");
const pkg = require_("@ksprtech/downword/package.json");
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
