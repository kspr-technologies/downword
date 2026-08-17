#!/usr/bin/env node
/**
 * Builds `packages/core` if — and only if — its `dist/` is missing or stale.
 *
 * The CLI consumes `downword` the way a published consumer does: through its
 * `exports` map, which points at `dist`. That is deliberate (it is the shape
 * that actually ships, and it means the CLI's tests exercise the real package
 * boundary), but it makes two commands order-dependent:
 *
 *   - `pnpm typecheck` — the CLI's types resolve to `dist/*.d.ts`; and
 *   - `pnpm test` — the built binary imports `downword` at runtime.
 *
 * CI runs `typecheck` and `test` **before** `build`, and a fresh clone has no
 * `dist` at all. Rather than reorder the pipeline for one package, the package
 * that has the dependency states it: both scripts call this first, and both are
 * then correct from a clean checkout, in any order, in CI and locally.
 *
 * `pnpm -r build` already builds core before the CLI (pnpm runs recursive
 * scripts in topological order), so in that path this is a stat walk and
 * nothing else.
 *
 * Usage: node scripts/ensure-core-build.mjs [--force]
 */

import { execFileSync } from "node:child_process";
import { existsSync, readdirSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const CLI_DIR = dirname(dirname(fileURLToPath(import.meta.url)));
const CORE_DIR = join(dirname(CLI_DIR), "core");

/**
 * Artifacts the CLI actually reaches for: the entry it imports, the two
 * optional entries it lazily imports, and the type files behind them.
 */
const REQUIRED = [
  "dist/index.js",
  "dist/index.d.ts",
  "dist/highlight.js",
  "dist/highlight.d.ts",
  "dist/images/node.js",
  "dist/images/node.d.ts",
  "dist/plugins/math.js",
  "dist/plugins/math.d.ts",
];

/** Inputs whose change should invalidate the build. */
const SOURCES = ["src", "package.json", "tsup.config.ts", "tsconfig.json"];

/** Newest mtime in a file or directory tree, or 0 when it does not exist. */
function newestMtime(path) {
  if (!existsSync(path)) return 0;
  const info = statSync(path);
  if (!info.isDirectory()) return info.mtimeMs;
  let newest = info.mtimeMs;
  for (const entry of readdirSync(path, { withFileTypes: true })) {
    newest = Math.max(newest, newestMtime(join(path, entry.name)));
  }
  return newest;
}

/** Why a build is needed, or `null` when it is not. */
function staleness() {
  const missing = REQUIRED.find((artifact) => !existsSync(join(CORE_DIR, artifact)));
  if (missing !== undefined) return `packages/core/${missing} is missing`;

  const newestSource = Math.max(...SOURCES.map((source) => newestMtime(join(CORE_DIR, source))));
  const oldestArtifact = Math.min(
    ...REQUIRED.map((artifact) => statSync(join(CORE_DIR, artifact)).mtimeMs),
  );
  return oldestArtifact < newestSource ? "packages/core/src is newer than its dist" : null;
}

const reason = process.argv.includes("--force") ? "--force" : staleness();
if (reason !== null) {
  process.stderr.write(`ensure-core-build: building packages/core (${reason})\n`);
  execFileSync("pnpm", ["run", "build"], { cwd: CORE_DIR, stdio: "inherit" });
}
