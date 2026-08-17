/**
 * Builds what the tests drive, once per run.
 *
 * Every test in this package spawns `dist/index.js` — the real binary, with the
 * real `downword` behind it — because a CLI that is only ever imported is a
 * library with an unusual signature. CI runs `pnpm test` before `pnpm build`,
 * and a fresh clone has neither `dist`, so the suite builds both itself:
 * `scripts/ensure-core-build.mjs` for the library, and `pnpm run build` here
 * when `src` is newer than the binary.
 *
 * Doing it in `globalSetup` rather than a `beforeAll` matters: vitest runs test
 * files in separate workers, and four workers racing four `tsup` invocations
 * over one `dist` directory is a flake waiting to happen.
 */

import { execFileSync } from "node:child_process";
import { existsSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const PKG_DIR = fileURLToPath(new URL("..", import.meta.url));
const BIN = join(PKG_DIR, "dist", "index.js");

/** Newest mtime under a directory tree. */
function newestMtime(dir: string): number {
  let newest = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    newest = Math.max(newest, entry.isDirectory() ? newestMtime(full) : statSync(full).mtimeMs);
  }
  return newest;
}

export default function setup(): void {
  execFileSync(process.execPath, [join(PKG_DIR, "scripts", "ensure-core-build.mjs")], {
    cwd: PKG_DIR,
    stdio: "inherit",
  });

  const stale = !existsSync(BIN) || statSync(BIN).mtimeMs < newestMtime(join(PKG_DIR, "src"));
  if (stale) {
    execFileSync("pnpm", ["run", "build"], { cwd: PKG_DIR, stdio: "inherit" });
  }
}
