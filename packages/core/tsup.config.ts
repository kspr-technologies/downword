import { readFileSync } from "node:fs";
import { defineConfig } from "tsup";

const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as {
  version: string;
};

/**
 * Five entry points, and the split between them is load-bearing.
 *
 * `index` is the one every consumer pays for, so it must stay free of anything
 * heavy: highlight.js (a megabyte, and an *optional* peer dependency), the Node
 * builtins the filesystem image loader needs, and the engines the math and
 * mermaid plugins will pull in when they land. Each of those lives behind its
 * own subpath instead, and `tests/bundle.test.ts` bundles `src/index.ts` for
 * the browser and fails if any of them reappears in the import graph.
 *
 * `splitting: false` keeps every entry self-contained. It costs a little
 * duplication between chunks and buys two things worth more: output a human can
 * read, and no shared chunk that a `require()` of one subpath drags in for
 * another. highlight.js's 63 `import()` calls survive it untouched, because
 * highlight.js is a peer dependency and therefore external - the *consumer's*
 * bundler is what turns them into chunks.
 */
export default defineConfig({
  entry: {
    index: "src/index.ts",
    highlight: "src/highlight/index.ts",
    "images/node": "src/images/node.ts",
    "plugins/math": "src/plugins/math.ts",
    "plugins/mermaid": "src/plugins/mermaid.ts",
  },
  format: ["esm", "cjs"],
  target: "es2022",
  platform: "neutral",
  // `platform: "neutral"` makes esbuild resolve nothing for free, so the Node
  // builtins `images/node` imports have to be declared external by hand.
  external: [/^node:/],
  // tsup rewrites `node:fs/promises` to `fs/promises` by default (a tsup 8
  // legacy, `removeNodeProtocol: true`). That is actively wrong here: the
  // prefix is what tells a bundler "this is a builtin, do not go looking for a
  // package called fs/promises", and it is the only spelling for newer builtins.
  // `images/node` is the sole entry that imports any of them.
  removeNodeProtocol: false,
  dts: true,
  sourcemap: true,
  treeshake: true,
  clean: true,
  splitting: false,
  define: {
    __DOWNWORD_VERSION__: JSON.stringify(pkg.version),
  },
});
