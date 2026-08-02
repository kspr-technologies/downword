import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

const PKG_DIR = fileURLToPath(new URL("..", import.meta.url));
const SRC_DIR = join(PKG_DIR, "src");
const BIN = join(PKG_DIR, "dist", "index.js");

const pkg = JSON.parse(readFileSync(join(PKG_DIR, "package.json"), "utf8")) as { version: string };

/** Newest mtime under a directory tree. */
function newestMtime(dir: string): number {
  let newest = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    newest = Math.max(newest, entry.isDirectory() ? newestMtime(full) : statSync(full).mtimeMs);
  }
  return newest;
}

/**
 * The CLI tests exercise the *built* binary, and CI runs `test` before `build`.
 * Build on demand (and rebuild when src is newer) so the test is self-contained
 * whatever order the scripts run in.
 */
beforeAll(() => {
  const needsBuild = !existsSync(BIN) || statSync(BIN).mtimeMs < newestMtime(SRC_DIR);
  if (needsBuild) {
    execFileSync("pnpm", ["run", "build"], { cwd: PKG_DIR, stdio: "inherit" });
  }
});

function runCli(args: readonly string[]): {
  status: number | null;
  stdout: string;
  stderr: string;
} {
  const result = spawnSync(process.execPath, [BIN, ...args], { encoding: "utf8" });
  return { status: result.status, stdout: result.stdout, stderr: result.stderr };
}

describe("downword CLI", () => {
  it("prints help and exits 0", () => {
    const { status, stdout, stderr } = runCli(["--help"]);

    expect(stderr).toBe("");
    expect(status).toBe(0);
    expect(stdout).toMatchSnapshot("--help");
  });

  it("prints help when invoked with no arguments", () => {
    const noArgs = runCli([]);
    const help = runCli(["--help"]);

    expect(noArgs.status).toBe(0);
    expect(noArgs.stdout).toBe(help.stdout);
  });

  it("prints the package version with --version", () => {
    const { status, stdout } = runCli(["--version"]);

    expect(status).toBe(0);
    expect(stdout.trim()).toBe(pkg.version);
  });

  it("rejects unknown options with EX_USAGE (64)", () => {
    const { status, stderr } = runCli(["--definitely-not-a-flag"]);

    expect(status).toBe(64);
    expect(stderr).toContain("Unknown option");
  });

  it("exits 70 for the not-yet-implemented conversion path", () => {
    const { status, stderr } = runCli(["README.md"]);

    expect(status).toBe(70);
    expect(stderr).toContain("not implemented");
  });

  it("is a real executable with a node shebang", () => {
    const contents = readFileSync(BIN, "utf8");
    expect(contents.startsWith("#!/usr/bin/env node\n")).toBe(true);

    expect(statSync(BIN).mode & 0o111).toBeGreaterThan(0);
  });
});
