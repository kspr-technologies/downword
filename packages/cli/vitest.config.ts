import { readFileSync } from "node:fs";
import { defineConfig } from "vitest/config";

const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as {
  version: string;
};

export default defineConfig({
  define: {
    __DOWNWORD_CLI_VERSION__: JSON.stringify(pkg.version),
  },
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
    // Every test file here drives the *built* binary, and CI runs `test` before
    // `build`. The global setup builds `packages/core` and this package on
    // demand — once per run, rather than once per worker.
    globalSetup: ["./tests/global-setup.ts"],
    testTimeout: 120_000,
    hookTimeout: 120_000,
  },
});
