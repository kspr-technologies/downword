/**
 * One markdown string in, one `.docx` out — and everything the CLI adds to
 * `convert()` on the way.
 *
 * ## Nothing here is imported until it is needed
 *
 * Every reference to `@ksprtech/downword` in this file is either an `import type` (erased
 * at build time) or an `await import(...)` inside the conversion. That is what
 * makes `downword --help` and `downword --version` answer without loading
 * `docx`, markdown-it or the converter, and it is what keeps the three optional
 * entry points — math, highlighting and anything with a DOM — off the startup
 * path of a run that does not ask for them. `tests/startup.test.ts` measures it
 * rather than trusting it, by listing every script V8 actually loaded.
 *
 * ## What the CLI adds
 *
 * | | |
 * | --- | --- |
 * | image resolver | `createNodeImageResolver`, rooted at the markdown file's own directory: local pictures work, `/etc/passwd` does not, and the network stays off unless `--remote-images` |
 * | mermaid | a fence scanner (`./mermaid.ts`) so a diagram the CLI cannot draw is named rather than silently left as code |
 * | maths | `@ksprtech/downword/plugins/math`, loaded only for `--math omml` / `--math image` |
 * | highlighting | `@ksprtech/downword/highlight`, loaded only for `--highlight` |
 *
 * ## Failure
 *
 * `convert()` throws only `DownwordError`. Each code is mapped to the exit code
 * that describes *whose* problem it is: a bad option is the command line's
 * (`EX_USAGE`), anything else is the software's (`EX_SOFTWARE`). Warnings are
 * not failures and never change the exit code — see `./diagnostics.ts`.
 */

import type {
  ConvertOptions,
  ConvertWarning,
  Highlighter,
  MarkdownItPlugin,
} from "@ksprtech/downword";
import type { MathWarning } from "@ksprtech/downword/plugins/math";

import type { Diagnostic } from "./diagnostics.js";
import { CliError, EXIT, describeError } from "./errors.js";
import { describeSkippedFence, scanMermaidFences, type MermaidFence } from "./mermaid.js";
import type { CliOptions } from "./options.js";

/** One document's worth of work. */
export interface ConversionRequest {
  /** The markdown source. */
  readonly markdown: string;
  /**
   * Directory that relative image paths resolve against, and (unless
   * `--allow-outside-images`) cannot escape. The input file's own directory, so
   * `![](./diagram.png)` next to `docs/notes.md` means `docs/diagram.png`.
   */
  readonly baseDir: string;
  readonly options: CliOptions;
  /** Receives every warning, in the order the pipeline found it. */
  readonly report: (diagnostic: Diagnostic) => void;
}

/** Longest snippet of TeX quoted back in a maths warning. */
const TEX_SNIPPET = 48;

/** Shortens a one-line snippet for a message. */
function snippet(text: string, limit: number): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length <= limit ? flat : `${flat.slice(0, limit - 1)}…`;
}

/**
 * Library advice, restated as something a shell user can act on.
 *
 * `downword`'s messages are written for whoever called `convert()`, so a
 * blocked image ends "pass `{ allowRemote: true }`" — which is exactly right for
 * a host application and useless at a prompt. The substitution is deliberately
 * literal: if the library rewords a message the replacement simply does not
 * fire and the original text is shown, rather than a mangled one.
 */
const FLAG_ADVICE: readonly (readonly [string, string])[] = [
  ["pass { allowRemote: true }", "pass --remote-images"],
  ["pass { allowOutsideBaseDir: true }", "pass --allow-outside-images"],
];

/** A library warning, in the CLI's own vocabulary. */
function toDiagnostic(warning: ConvertWarning): Diagnostic {
  const line = warning.stage === "parse" ? warning.line : null;
  // Image diagnostics know which source failed; most messages name it already,
  // and the ones that talk about the host rather than the URL do not.
  const located =
    warning.stage === "image" && !warning.message.includes(warning.src)
      ? `${warning.message} (${snippet(warning.src, 80)})`
      : warning.message;
  const message = FLAG_ADVICE.reduce((text, [library, cli]) => text.replace(library, cli), located);
  return { severity: warning.severity, label: `${warning.stage}/${warning.code}`, message, line };
}

/** A maths warning, likewise. It has no line, so it quotes the equation instead. */
function toMathDiagnostic(warning: MathWarning): Diagnostic {
  const tex = warning.tex === "" ? "" : ` (in ${snippet(warning.tex, TEX_SNIPPET)})`;
  return {
    severity: warning.severity,
    label: `math/${warning.code}`,
    message: `${warning.message}${tex}`,
    line: null,
  };
}

/** Loads the optional highlighter, degrading to plain code blocks if it cannot. */
async function loadHighlighter(report: ConversionRequest["report"]): Promise<Highlighter | null> {
  try {
    const { createHighlighter } = await import("@ksprtech/downword/highlight");
    return createHighlighter();
  } catch (error: unknown) {
    report({
      severity: "error",
      label: "highlight/unavailable",
      message:
        `--highlight could not load @ksprtech/downword/highlight (${describeError(error)}); ` +
        `code blocks stay plain. Install the optional peer: npm i highlight.js`,
      line: null,
    });
    return null;
  }
}

/** Maps a `DownwordError` code onto the exit code that describes it. */
function conversionFailure(error: unknown): CliError {
  const code: unknown =
    typeof error === "object" && error !== null ? (error as { code?: unknown }).code : undefined;
  const usage = code === "invalid-options" || code === "invalid-input";
  return new CliError(usage ? EXIT.usage : EXIT.software, describeError(error), { cause: error });
}

/**
 * Converts one document.
 *
 * @param request - See {@link ConversionRequest}.
 * @returns The `.docx` bytes.
 * @throws {CliError} and nothing else.
 */
export async function convertMarkdown(request: ConversionRequest): Promise<Uint8Array> {
  const { markdown, options, report } = request;

  const fences: MermaidFence[] = [];
  const plugins: MarkdownItPlugin[] = [
    scanMermaidFences((fence) => {
      fences.push(fence);
    }),
  ];

  const [{ convert }, { createNodeImageResolver }] = await Promise.all([
    import("@ksprtech/downword"),
    import("@ksprtech/downword/images/node"),
  ]);
  const highlighter = options.highlight ? await loadHighlighter(report) : null;

  const convertOptions: ConvertOptions = {
    theme: options.theme,
    pageSize: options.pageSize,
    toc: options.toc,
    pageNumbers: options.pageNumbers,
    footnotes: options.footnotes,
    metadata: {
      title: options.title ?? undefined,
      author: options.author ?? undefined,
    },
    plugins,
    imageResolver: createNodeImageResolver({
      baseDir: request.baseDir,
      allowOutsideBaseDir: options.allowOutsideImages,
      // Left at its default (10 MiB / 10 s), which is the browser build's cap
      // too: one number, one place, so the two runtimes cannot disagree about
      // how big "too big" is.
      allowRemote: options.remoteImages,
    }),
    onWarning: (warning) => {
      report(toDiagnostic(warning));
    },
    ...(highlighter === null ? {} : { highlighter }),
  };

  let bytes: Uint8Array;
  try {
    if (options.math === "off") {
      bytes = await convert(markdown, convertOptions);
    } else {
      const { convertWithMath } = await import("@ksprtech/downword/plugins/math");
      bytes = await convertWithMath(markdown, {
        ...convertOptions,
        math: {
          math: options.math,
          onWarning: (warning) => {
            report(toMathDiagnostic(warning));
          },
        },
      });
    }
  } catch (error: unknown) {
    throw conversionFailure(error);
  }

  // Reported after the fact because the scan happens during the parse, inside
  // convert(). One line per diagram, naming it: see `./mermaid.ts`.
  fences.forEach((fence, index) => {
    report({
      severity: "notice",
      always: true,
      label: "mermaid/no-dom",
      message: describeSkippedFence(fence, index),
      line: fence.line,
    });
  });

  return bytes;
}
