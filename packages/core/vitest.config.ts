import { readFileSync } from "node:fs";
import { defineConfig } from "vitest/config";

const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as {
  version: string;
};

export default defineConfig({
  define: {
    __DOWNWORD_VERSION__: JSON.stringify(pkg.version),
  },
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
    // Golden tests write .docx fixtures to disk; keep the run deterministic.
    sequence: {
      shuffle: false,
    },
  },
});
