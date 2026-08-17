/**
 * `--help`, and the one-line nudge printed after a usage error.
 *
 * The text is a snapshot test (`tests/cli.test.ts`), which is the point: help
 * output is the CLI's actual documentation, and a flag that changes behaviour
 * without changing this file will fail the build.
 *
 * Two things it says out loud rather than leaving to be discovered:
 *
 *  - **the network is off**, and `--no-remote-images` names the default rather
 *    than only existing as the inverse of an opt-in nobody read; and
 *  - **mermaid needs a browser**, so a fence that comes back as a code block is
 *    a documented outcome rather than a bug report.
 *
 * Kept under 80 columns so it survives a narrow terminal without reflowing.
 */

/** The `--help` text. Ends with a newline; write it verbatim. */
export const HELP = `downword — Markdown to Word (.docx)

Usage:
  downword [options] <input.md>...
  downword [options] "docs/**/*.md" --outdir build/
  cat in.md | downword > out.docx

Inputs:
  <input.md>...          One or more markdown files. Quote glob patterns
                         ("docs/*.md") so downword expands them rather
                         than the shell. With no input at all, markdown
                         is read from stdin.

Output:
  -o, --output <file>    Write the .docx here; "-" means stdout.
                         Default: next to each input (in.md -> in.docx),
                         or stdout when reading stdin.
      --outdir <dir>     Batch mode: one .docx per input, written into
                         <dir> (created if missing). Not with --output.

Document:
      --theme <name>     default | github | academic | academic-double
                         | print                        (default: default)
      --page-size <size> A4 | Letter                    (default: A4)
      --title <text>     Document title, as Word shows it under File > Info
      --author <name>    Document author, likewise
      --toc              Insert a Word table-of-contents field. It is
                         empty until the reader updates it — Word offers
                         to on open; most other readers never do.
      --page-numbers     Centre a page number in every page's footer.
      --no-footnotes     Leave [^1] as literal text instead of turning it
                         into a real Word footnote.  (default: footnotes on)
      --math <mode>      omml | image | off                (default: off)
                           omml   $x^2$ and $$…$$ become native, editable
                                  Word equations. Needs the optional peer
                                  packages: npm i temml mathml2omml
                           image  no rasteriser ships with the CLI, so this
                                  reports a notice and falls back to omml
                           off    a dollar sign is just a dollar sign
      --highlight        Colour fenced code blocks. Needs the optional
                         peer package: npm i highlight.js

Images (the network is off unless you turn it on):
      --no-remote-images The default, stated explicitly: http(s) image
                         sources are refused and no request is made.
      --remote-images    Allow them. Each fetch is capped at 10 MiB and
                         10 s — the same caps as the browser build — and
                         hosts resolving to private or loopback addresses
                         are refused.
      --allow-outside-images
                         Let ![](../logo.png) read files outside the input
                         file's own directory.            (default: refused)

Diagnostics:
  -q, --quiet            Print nothing but fatal errors.
      --verbose          Print every warning, plus a line per file.
                         By default, warnings that cost the document
                         content are printed and notices are counted.
  -v, --version          Print the version and exit.
  -h, --help             Print this help and exit.

Mermaid diagrams need a browser DOM to measure text, so the CLI cannot
draw them: a \`\`\`mermaid fence is left as a code block and named on
stderr. Pre-render it and reference the picture with ![](diagram.png).

Exit codes: 0 ok · 64 bad usage · 66 input not found · 70 conversion
failed · 73 cannot write output.

Docs: https://github.com/kspr-technologies/downword
`;

/** Printed under a usage error, so the reader is one command from the answer. */
export const TRY_HELP = "Try 'downword --help' for the full list of options.";
