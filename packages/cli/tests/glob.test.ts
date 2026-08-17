/**
 * The two glob implementations, held to the same answer.
 *
 * `fs.promises.glob` exists on Node 22 and not on Node 20, and CI runs both, so
 * the risk this file addresses is not "does the matcher work" but "are the two
 * legs of CI running the same program". Every case is asserted against the
 * built-in matcher *and*, where the platform has one, against Node's — with a
 * final test that the two agree over the whole fixture tree.
 */

import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { expandBraces, expandPattern, globFallback, hasMagic, nativeGlob } from "../src/glob.js";

let root: string;

/** A tree with the shapes that separate a real matcher from a regex. */
beforeAll(() => {
  root = mkdtempSync(join(tmpdir(), "downword-glob-"));
  mkdirSync(join(root, "docs", "sub", "deep"), { recursive: true });
  mkdirSync(join(root, "docs", "folder.md"), { recursive: true });
  mkdirSync(join(root, ".hidden"), { recursive: true });
  for (const path of [
    "readme.md",
    "docs/a.md",
    "docs/b.md",
    "docs/notes.txt",
    "docs/.secret.md",
    "docs/sub/c.md",
    "docs/sub/deep/d.md",
    ".hidden/h.md",
  ]) {
    writeFileSync(join(root, path), `# ${path}\n`);
  }
});

afterAll(() => {
  rmSync(root, { recursive: true, force: true });
});

/** Both implementations, or just the fallback where Node has no glob. */
async function bothWays(pattern: string): Promise<{ fallback: string[]; native: string[] | null }> {
  const glob = nativeGlob();
  return {
    fallback: await globFallback(pattern, root),
    // Node's glob resolves against process.cwd(); point it at the fixture tree.
    native: glob === null ? null : await inCwd(root, () => expandPattern(pattern, root, glob)),
  };
}

/** Runs `body` with `process.cwd()` pointed at `dir`. */
async function inCwd<T>(dir: string, body: () => Promise<T>): Promise<T> {
  const previous = process.cwd();
  process.chdir(dir);
  try {
    return await body();
  } finally {
    process.chdir(previous);
  }
}

describe("hasMagic", () => {
  it("tells a pattern from a path", () => {
    expect(hasMagic("docs/*.md")).toBe(true);
    expect(hasMagic("docs/**/x.md")).toBe(true);
    expect(hasMagic("a?.md")).toBe(true);
    expect(hasMagic("a[12].md")).toBe(true);
    expect(hasMagic("{a,b}.md")).toBe(true);
    expect(hasMagic("docs/notes.md")).toBe(false);
    expect(hasMagic("/absolute/path.md")).toBe(false);
  });
});

describe("expandBraces", () => {
  it("expands, nests, and leaves unbalanced braces alone", () => {
    expect(expandBraces("a.md")).toEqual(["a.md"]);
    expect(expandBraces("{a,b}.md")).toEqual(["a.md", "b.md"]);
    expect(expandBraces("docs/{a,b}/{c,d}.md")).toEqual([
      "docs/a/c.md",
      "docs/a/d.md",
      "docs/b/c.md",
      "docs/b/d.md",
    ]);
    expect(expandBraces("{a,{b,c}}.md")).toEqual(["a.md", "b.md", "c.md"]);
    expect(expandBraces("a{b.md")).toEqual(["a{b.md"]);
  });
});

describe("matching", () => {
  it.each([
    ["docs/*.md", ["docs/a.md", "docs/b.md"]],
    ["docs/?.md", ["docs/a.md", "docs/b.md"]],
    ["docs/[ab].md", ["docs/a.md", "docs/b.md"]],
    ["docs/[!a].md", ["docs/b.md"]],
    ["docs/{a,notes}.{md,txt}", ["docs/a.md", "docs/notes.txt"]],
    ["docs/**/*.md", ["docs/a.md", "docs/b.md", "docs/sub/c.md", "docs/sub/deep/d.md"]],
    ["**/*.md", ["docs/a.md", "docs/b.md", "docs/sub/c.md", "docs/sub/deep/d.md", "readme.md"]],
    ["docs/sub/**/d.md", ["docs/sub/deep/d.md"]],
  ])("%s", async (pattern, expected) => {
    const { fallback, native } = await bothWays(pattern);

    expect(fallback).toEqual(expected);
    if (native !== null) expect(native).toEqual(expected);
  });

  it("never returns a directory, however well it matches", async () => {
    // `docs/folder.md` is a directory whose name ends in .md.
    const { fallback, native } = await bothWays("docs/*.md");

    expect(fallback).not.toContain(join("docs", "folder.md"));
    if (native !== null) expect(native).not.toContain(join("docs", "folder.md"));
  });

  it("skips dotfiles unless the pattern spells them out", async () => {
    expect((await bothWays("docs/*.md")).fallback).not.toContain(join("docs", ".secret.md"));
    expect((await bothWays("**/*.md")).fallback).not.toContain(join(".hidden", "h.md"));
    expect(await globFallback("docs/.*.md", root)).toEqual([join("docs", ".secret.md")]);
  });

  it("matches nothing, rather than everything, when nothing matches", async () => {
    const { fallback, native } = await bothWays("docs/*.rst");

    expect(fallback).toEqual([]);
    if (native !== null) expect(native).toEqual([]);
  });

  it("agrees with Node's own glob across the whole tree", async () => {
    const glob = nativeGlob();
    if (glob === null) return; // Node 20: there is only one implementation.

    for (const pattern of ["**/*", "**/*.md", "docs/**", "docs/*", "*"]) {
      const native = await inCwd(root, () => expandPattern(pattern, root, glob));
      expect({ pattern, files: await globFallback(pattern, root) }).toEqual({
        pattern,
        files: native,
      });
    }
  });
});
