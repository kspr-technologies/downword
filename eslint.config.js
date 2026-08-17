// @ts-check
import js from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    // Keep this first and standalone: a config object with only `ignores`
    // is a global ignore list in flat config.
    ignores: [
      "**/node_modules/**",
      "**/dist/**",
      "**/coverage/**",
      "**/tests/__fixtures__/**",
      "**/tests/**/__snapshots__/**",
      "**/*.tgz",
      "examples/*/out/**",
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/consistent-type-imports": [
        "error",
        { prefer: "type-imports", fixStyle: "inline-type-imports" },
      ],
      "@typescript-eslint/no-import-type-side-effects": "error",
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
        },
      ],
      eqeqeq: ["error", "always", { null: "ignore" }],
      "no-console": "off",
    },
  },
  {
    // The CommonJS example exists precisely to exercise `require()`.
    files: ["**/*.cjs"],
    rules: {
      "@typescript-eslint/no-require-imports": "off",
    },
  },
  {
    // Node-flavoured scripts and configs.
    files: ["**/scripts/**/*.mjs", "**/*.config.{js,ts,mjs}", "examples/**/*.{mjs,cjs,js}"],
    languageOptions: {
      globals: {
        console: "readonly",
        process: "readonly",
        __dirname: "readonly",
        __filename: "readonly",
        require: "readonly",
        module: "writable",
        Buffer: "readonly",
        URL: "readonly",
      },
    },
  },
);
