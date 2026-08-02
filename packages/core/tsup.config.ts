import { readFileSync } from "node:fs";
import { defineConfig } from "tsup";

const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as {
  version: string;
};

export default defineConfig({
  entry: {
    index: "src/index.ts",
    "plugins/math": "src/plugins/math.ts",
    "plugins/mermaid": "src/plugins/mermaid.ts",
  },
  format: ["esm", "cjs"],
  target: "es2022",
  platform: "neutral",
  dts: true,
  sourcemap: true,
  treeshake: true,
  clean: true,
  splitting: false,
  define: {
    __DOWNWORD_VERSION__: JSON.stringify(pkg.version),
  },
});
