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
    // The CLI test builds the binary on demand (see tests/cli.test.ts).
    testTimeout: 120_000,
    hookTimeout: 120_000,
  },
});
