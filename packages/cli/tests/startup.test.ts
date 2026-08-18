/**
 * What the CLI actually loads, measured rather than asserted in a comment.
 *
 * The library is split into a small core plus four optional entry points, three
 * of which are expensive or browser-only: `@ksprtech/downword/plugins/mermaid` needs a
 * DOM and pulls a ~500 kB dependency, `@ksprtech/downword/plugins/math` pulls two, and
 * `@ksprtech/downword/highlight` pulls highlight.js. A CLI that imported them "just in
 * case" would pay for all of it on every run, including `--help`.
 *
 * Each test runs the real binary under `NODE_V8_COVERAGE`, which makes V8 write
 * out one record per script it compiled — the loaded module graph, from the
 * runtime, with nothing inferred from the bundle. The assertions are then
 * exact: not "we did not import mermaid" but "no script whose path mentions
 * mermaid was ever compiled in that process".
 */

import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import { afterAll, describe, expect, it } from "vitest";

import { BIN, FIXTURES, PKG_DIR, outPath } from "./helpers/cli.js";

/**
 * `file:///…/packages/core/dist` — the library's build output.
 *
 * Matched as a URL prefix rather than by the substring "downword", because the
 * repository directory is itself called `downword` and every path in this tree
 * therefore contains it. A test that always passes is worse than no test.
 */
const CORE_DIST = pathToFileURL(join(PKG_DIR, "..", "core", "dist")).href;

const scratch: string[] = [];

afterAll(() => {
  for (const dir of scratch) rmSync(dir, { recursive: true, force: true });
});

/** Every script URL V8 compiled while the CLI ran, excluding Node's internals. */
function loadedScripts(args: readonly string[]): string[] {
  const dir = mkdtempSync(join(tmpdir(), "downword-coverage-"));
  scratch.push(dir);

  const result = spawnSync(process.execPath, [BIN, ...args], {
    cwd: PKG_DIR,
    input: "",
    env: { ...process.env, NODE_V8_COVERAGE: dir },
    encoding: "buffer",
    maxBuffer: 32 * 1024 * 1024,
  });
  expect(result.status, (result.stderr ?? Buffer.alloc(0)).toString("utf8")).toBe(0);

  const urls: string[] = [];
  for (const name of readdirSync(dir)) {
    const report = JSON.parse(readFileSync(join(dir, name), "utf8")) as {
      result: { url: string }[];
    };
    for (const script of report.result) {
      if (script.url.startsWith("file:")) urls.push(script.url);
    }
  }
  expect(urls.length).toBeGreaterThan(0);
  return urls;
}

/** Scripts whose path contains `needle`, case-insensitively. */
function matching(urls: readonly string[], needle: string): string[] {
  return urls.filter((url) => url.toLowerCase().includes(needle.toLowerCase()));
}

/** Scripts loaded out of `packages/core/dist`. */
function fromLibrary(urls: readonly string[]): string[] {
  return urls.filter((url) => url.startsWith(`${CORE_DIST}/`));
}

describe("--help costs nothing", () => {
  it("answers without loading the library at all", () => {
    const urls = loadedScripts(["--help"]);

    expect(fromLibrary(urls)).toEqual([]);
    expect(matching(urls, "node_modules/docx")).toEqual([]);
    expect(matching(urls, "node_modules/markdown-it")).toEqual([]);
    // It did run the binary, though.
    expect(matching(urls, "cli/dist/index.js").length).toBe(1);
  });

  it("so does --version", () => {
    expect(fromLibrary(loadedScripts(["--version"]))).toEqual([]);
  });
});

describe("a real conversion loads the core and nothing browser-only", () => {
  const urls = (): string[] =>
    loadedScripts([join(FIXTURES, "mermaid.md"), "-o", outPath("cli-startup.docx")]);

  it("loads the library's main entry and the Node image resolver, and only those", () => {
    const loaded = fromLibrary(urls());

    expect(loaded.sort()).toEqual([`${CORE_DIST}/images/node.js`, `${CORE_DIST}/index.js`]);
  });

  it("never loads mermaid, in any form, even for a document full of diagrams", () => {
    expect(matching(urls(), "mermaid")).toEqual([]);
  });

  it("never loads the maths entry, temml or mathml2omml", () => {
    const loaded = urls();

    expect(matching(loaded, "plugins/math")).toEqual([]);
    expect(matching(loaded, "temml")).toEqual([]);
    expect(matching(loaded, "mathml2omml")).toEqual([]);
  });

  it("never loads highlight.js", () => {
    const loaded = urls();

    expect(matching(loaded, "dist/highlight")).toEqual([]);
    expect(matching(loaded, "node_modules/highlight.js")).toEqual([]);
  });
});

describe("the optional entries load when, and only when, they are asked for", () => {
  it("--math omml loads the maths entry", () => {
    const loaded = loadedScripts([
      join(FIXTURES, "math.md"),
      "-o",
      outPath("cli-startup-math.docx"),
      "--math",
      "omml",
    ]);

    expect(matching(loaded, "plugins/math").length).toBeGreaterThan(0);
    // …and still no mermaid.
    expect(matching(loaded, "mermaid")).toEqual([]);
  });

  it("--highlight loads the highlighting entry", () => {
    const loaded = loadedScripts([
      join(FIXTURES, "kitchen-sink.md"),
      "-o",
      outPath("cli-startup-highlight.docx"),
      "--highlight",
    ]);

    expect(matching(loaded, "dist/highlight").length).toBeGreaterThan(0);
  });
});

describe("the built binary's static imports", () => {
  it("mention no browser-only entry point", () => {
    const source = readFileSync(BIN, "utf8");
    const statics = [...source.matchAll(/^import[^\n]*from\s*"([^"]+)"/gm)].map(
      (match) => match[1] ?? "",
    );

    expect(statics).not.toContain("@ksprtech/downword");
    expect(statics).not.toContain("@ksprtech/downword/plugins/mermaid");
    expect(statics).not.toContain("@ksprtech/downword/plugins/math");
    expect(statics).not.toContain("@ksprtech/downword/highlight");
    // Everything it does import statically is a Node builtin.
    expect(statics.filter((specifier) => !specifier.startsWith("node:"))).toEqual([]);
  });
});
