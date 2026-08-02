import { parseArgs } from "node:util";

// NOTE: `downword` (packages/core) is already declared as a `workspace:*`
// dependency, but is deliberately not imported yet: this is the T0 scaffold and
// the CLI only handles --help/--version. Wiring it up is tracked as part of the
// renderer work. Keeping the import out of the tree means `pnpm typecheck` does
// not require `packages/core/dist` to exist first.

/** Version of this package, injected at build time. See `types/globals.d.ts`. */
export const CLI_VERSION: string = __DOWNWORD_CLI_VERSION__;

const HELP = `downword - Markdown to Word (.docx)

Usage:
  downword [options] <input.md>

Options:
  -o, --output <file>   Write the .docx here (default: alongside the input)
  -v, --version         Print the version and exit
  -h, --help            Print this help and exit

Status:
  This is the T0 scaffold. Conversion is not wired into the CLI yet;
  only --help and --version do anything today.

Docs:  https://github.com/kspr-technologies/downword
`;

/**
 * Runs the CLI.
 *
 * @param argv - arguments after `node <script>`, i.e. `process.argv.slice(2)`.
 * @returns the process exit code.
 */
export function run(argv: readonly string[]): number {
  let parsed;
  try {
    parsed = parseArgs({
      args: [...argv],
      options: {
        output: { type: "string", short: "o" },
        version: { type: "boolean", short: "v", default: false },
        help: { type: "boolean", short: "h", default: false },
      },
      allowPositionals: true,
      strict: true,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(`downword: ${message}\n\n${HELP}`);
    return 64; // EX_USAGE
  }

  if (parsed.values.version === true) {
    process.stdout.write(`${CLI_VERSION}\n`);
    return 0;
  }

  if (parsed.values.help === true || parsed.positionals.length === 0) {
    process.stdout.write(HELP);
    return 0;
  }

  process.stderr.write(
    "downword: conversion is not implemented in the CLI yet (T0 scaffold).\n" +
      "Track it at https://github.com/kspr-technologies/downword/issues\n",
  );
  return 70; // EX_SOFTWARE
}

process.exitCode = run(process.argv.slice(2));
