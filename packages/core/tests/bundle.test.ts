import { fileURLToPath } from "node:url";

import { build, type BuildOptions, type Metafile } from "esbuild";
import { describe, expect, it } from "vitest";

/**
 * Bundle discipline, proven rather than asserted.
 *
 * `@ksprtech/downword`'s main entry has one job beyond converting markdown: it must
 * cheap and it must stay browser-safe. Both are one careless `import` away from
 * being untrue, and neither shows up in a type error or a failing unit test —
 * the package would keep working perfectly while quietly costing every consumer
 * a megabyte, or breaking every consumer who bundles for a browser.
 *
 * So this file bundles the entry for real, with esbuild, and reads the
 * `metafile`'s input list. If a static `import "highlight.js"` appears anywhere
 * in the graph, this test fails; if `src/images/node.ts` (the one module that
 * touches `node:fs`) is reachable, this test fails.
 *
 * ### Why the *source* and not `dist/`
 *
 * A regression is introduced in `src/`, and the build can only ever reflect it,
 * so checking `src/` catches it at `pnpm test` — before `pnpm build` has even
 * run, which is the order CI uses. The shipped artifact is covered separately:
 * `pnpm size` bundles `dist/index.js` against a 150 kB budget (highlight.js
 * alone is ~7x that), and `scripts/pack-smoke.mjs` imports the real tarball.
 */

const PACKAGE_DIR = fileURLToPath(new URL("..", import.meta.url));

/**
 * Packages that must never be in the main entry's graph.
 *
 * Combined, they are several megabytes. Each is reachable only from a subpath
 * export (`@ksprtech/downword/highlight`) or is a dependency of a plugin entry point that
 * does not exist yet (`@ksprtech/downword/plugins/math`, `.../mermaid`) — listing the
 * future ones now means the guard is already in place when they land.
 */
const FORBIDDEN_PACKAGES = ["highlight.js", "katex", "temml", "mathml2omml", "mermaid"] as const;

/** Node builtins, with and without the `node:` prefix. */
const NODE_BUILTIN =
  /^(?:node:|(?:assert|buffer|child_process|crypto|dns|fs|http|https|module|net|os|path|process|stream|tls|url|util|worker_threads|zlib)(?:\/|$))/;

interface Bundled {
  readonly inputs: readonly string[];
  readonly externals: readonly string[];
  readonly text: string;
  readonly bytes: number;
}

async function bundle(entry: string, options: BuildOptions = {}): Promise<Bundled> {
  const result = await build({
    absWorkingDir: PACKAGE_DIR,
    entryPoints: [entry],
    bundle: true,
    format: "esm",
    target: "es2022",
    platform: "browser",
    write: false,
    metafile: true,
    logLevel: "silent",
    define: { __DOWNWORD_VERSION__: '"0.0.0-test"' },
    ...options,
  });

  const metafile = result.metafile as Metafile;
  const output = Object.values(metafile.outputs)[0];
  if (output === undefined) throw new Error(`esbuild produced no output for ${entry}`);

  return {
    inputs: Object.keys(metafile.inputs),
    externals: output.imports.filter((entry_) => entry_.external).map((entry_) => entry_.path),
    text: result.outputFiles?.[0]?.text ?? "",
    bytes: output.bytes,
  };
}

/**
 * The npm package a bundled input belongs to, or `null` for our own source.
 *
 * The *last* `node_modules/` wins: pnpm's real paths look like
 * `node_modules/.pnpm/docx@9.7.1/node_modules/docx/dist/index.mjs`, and the
 * first match would name every dependency `.pnpm`.
 */
function packageOf(input: string): string | null {
  const marker = input.lastIndexOf("node_modules/");
  if (marker === -1) return null;
  const rest = input.slice(marker + "node_modules/".length).split("/");
  const [first, second] = rest;
  if (first === undefined) return null;
  return first.startsWith("@") && second !== undefined ? `${first}/${second}` : first;
}

describe("main entry: no heavy static dependencies", () => {
  it("does not pull in highlight.js, katex, temml, mathml2omml or mermaid", async () => {
    const { inputs } = await bundle("src/index.ts");
    const packages = new Set(inputs.map(packageOf).filter((name) => name !== null));

    for (const forbidden of FORBIDDEN_PACKAGES) {
      expect(packages, `${forbidden} must not be in the main entry's import graph`).not.toContain(
        forbidden,
      );
    }
  });

  it("pulls in exactly the three runtime dependencies it declares", async () => {
    const { inputs } = await bundle("src/index.ts");
    const packages = [...new Set(inputs.map(packageOf).filter((name) => name !== null))].sort();

    // markdown-it's own dependencies (mdurl, uc.micro, entities, linkify-it,
    // punycode.js) come along with it; a *new* name here means a new dependency
    // landed in the main entry and should be a deliberate decision.
    expect(packages).toEqual([
      "docx",
      "entities",
      "linkify-it",
      "markdown-it",
      "markdown-it-footnote",
      "mdurl",
      "punycode.js",
      "uc.micro",
    ]);
  });

  it("keeps the Node-only image resolver and the plugin stubs out of the graph", async () => {
    const { inputs } = await bundle("src/index.ts");
    const own = inputs.filter((input) => input.startsWith("src/"));

    expect(own).not.toContain("src/images/node.ts");
    expect(own.filter((input) => input.startsWith("src/plugins/"))).toEqual([]);
  });

  it("touches only the one highlight module that holds no engine code", async () => {
    const { inputs } = await bundle("src/index.ts");

    // src/highlight/palette.ts is a table of colours with zero imports; it is
    // in the graph because THEMES.print reuses the measured AAA-contrast
    // palette rather than duplicating it. Everything else under src/highlight/
    // - the loader table, the engine adapter, the emitter - belongs to the
    // `@ksprtech/downword/highlight` subpath and must stay there.
    expect(inputs.filter((input) => input.startsWith("src/highlight/"))).toEqual([
      "src/highlight/palette.ts",
    ]);
  });
});

describe("main entry: browser safety", () => {
  it("bundles for the browser with no errors and no externals", async () => {
    const { externals } = await bundle("src/index.ts");

    // With platform: "browser" esbuild has no builtin shims, so an unresolved
    // `node:fs` is a hard build error - a successful build with an empty
    // external list *is* the proof that nothing Node-only is reachable.
    expect(externals).toEqual([]);
  });

  it("imports no Node builtin, prefixed or bare", async () => {
    // `packages: "external"` bundles nothing, so *every* bare specifier the
    // entry reaches shows up in the external list - which is where a
    // `node:fs` would be, and is a check that does not depend on esbuild
    // choosing to fail the build.
    const { externals } = await bundle("src/index.ts", { packages: "external" });

    expect(externals.filter((specifier) => NODE_BUILTIN.test(specifier))).toEqual([]);
    // esbuild lists one entry per import statement, hence the dedupe.
    expect([...new Set(externals)].sort()).toEqual(["docx", "markdown-it", "markdown-it-footnote"]);
  });

  it("negative control: the Node image resolver genuinely cannot be bundled for a browser", async () => {
    // If this ever passes, the assertions above have stopped meaning anything.
    await expect(bundle("src/images/node.ts")).rejects.toThrow(/Could not resolve "node:/);
  });
});

describe("@ksprtech/downword/highlight: highlight.js is loaded lazily or not at all", () => {
  it("has no static import of highlight.js, only dynamic ones", async () => {
    // `packages: "external"` reproduces what tsup publishes: highlight.js is an
    // optional peer dependency, so its specifiers survive into the output and
    // the consumer's bundler decides how to chunk them.
    const { text } = await bundle("src/highlight/index.ts", { packages: "external" });

    const staticImports = [...text.matchAll(/(?:^|\n)import\s[^;]*?from\s*["']([^"']+)["']/g)].map(
      (match) => match[1],
    );
    const dynamicImports = [...text.matchAll(/import\(\s*["']([^"']+)["']\s*\)/g)].map(
      (match) => match[1] as string,
    );

    expect(staticImports).toEqual([]);
    expect(dynamicImports).toContain("highlight.js/lib/core");
    expect(
      dynamicImports.filter((path) => path.startsWith("highlight.js/lib/languages/")).length,
    ).toBeGreaterThan(50);
  });

  it("costs no engine code to import: the eager bundle is a few kilobytes of tables", async () => {
    const { bytes } = await bundle("src/highlight/index.ts", { packages: "external" });

    // highlight.js's core alone is ~130 kB unminified. A regression that made
    // the engine eager would blow straight past this.
    expect(bytes).toBeLessThan(64 * 1024);
  });
});
