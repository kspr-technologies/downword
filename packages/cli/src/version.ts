/**
 * The version string, injected at build time.
 *
 * It lives in its own module so that `--version` costs one tiny import rather
 * than pulling in the entry point's whole graph, and so both `tsup` (`define`)
 * and `vitest` (`define`) have exactly one place to substitute. See
 * `types/globals.d.ts`.
 */

/** This package's version, from `packages/cli/package.json`. */
export const CLI_VERSION: string = __DOWNWORD_CLI_VERSION__;
