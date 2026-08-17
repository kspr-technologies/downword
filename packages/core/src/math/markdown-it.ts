/**
 * The parse half: `$…$` and `$$…$$` into `math_inline` / `math_block` tokens.
 *
 * `downword`'s parser already maps those two token types onto `MathInlineNode`
 * and `MathBlockNode` — it just has no opinion about who produces them, so that
 * a host already using `@vscode/markdown-it-katex` can keep it. This is the
 * built-in producer, and it is deliberately tiny: no engine, no dependency, ~1
 * kB, and it is the only part of the math plugin that runs during parsing.
 *
 * ## The delimiter rules, and why they are fussy
 *
 * A dollar sign is also a currency symbol, and `it costs $5 and $10` must not
 * become an equation. The conventions below are the ones pandoc and
 * `markdown-it-katex` settled on, and they get that case right twice over:
 *
 * - an opening `$` is not followed by whitespace — `$ x$` is not maths;
 * - a closing `$` is not preceded by whitespace — `$x $` is not maths;
 * - a closing `$` is not followed by a digit — which is what rules out
 *   `$5 and $10`;
 * - `\$` is an escaped dollar and never a delimiter, however many backslashes
 *   precede it;
 * - the content is never empty.
 *
 * Everything that fails those tests is left alone for the ordinary text rule to
 * pick up, so a document full of prices parses exactly as it did before.
 *
 * ## Display maths
 *
 * `$$` opens a block whether or not it is on a line of its own, and the block
 * runs to the next line ending in `$$`. An **unterminated** `$$` closes at the
 * end of the containing block rather than swallowing the rest of the document
 * into an error — the same forgiving policy the parser applies to an unclosed
 * code fence.
 */

import type { MarkdownIt, StateBlock, StateInline } from "markdown-it";

import type { MarkdownItPlugin } from "../parse/index.js";

const DOLLAR = 0x24;
const BACKSLASH = 0x5c;
const SPACE = 0x20;
const TAB = 0x09;
const NEWLINE = 0x0a;
const ZERO = 0x30;
const NINE = 0x39;

function isSpace(code: number): boolean {
  return code === SPACE || code === TAB || code === NEWLINE;
}

function isDigit(code: number): boolean {
  return code >= ZERO && code <= NINE;
}

/**
 * The next unescaped occurrence of `marker` at or after `from`, or `-1`.
 *
 * "Unescaped" counts the backslashes immediately before it: an even number
 * means the run is itself escaped and the delimiter is real, an odd number
 * means the delimiter is not.
 */
function findClosing(src: string, from: number, max: number, marker: string): number {
  let at = src.indexOf(marker, from);
  while (at !== -1 && at + marker.length <= max) {
    let back = at - 1;
    while (back >= 0 && src.charCodeAt(back) === BACKSLASH) back -= 1;
    if ((at - back) % 2 === 1) return at;
    at = src.indexOf(marker, at + 1);
  }
  return -1;
}

/** One `$…$` or `$$…$$` inside a paragraph. */
function inlineRule(state: StateInline, silent: boolean): boolean {
  const { src, pos, posMax } = state;
  if (src.charCodeAt(pos) !== DOLLAR) return false;

  const double = src.charCodeAt(pos + 1) === DOLLAR;
  const marker = double ? "$$" : "$";
  const start = pos + marker.length;
  if (start >= posMax) return false;
  // `$$$` and `$ x$` open nothing.
  if (isSpace(src.charCodeAt(start)) || src.charCodeAt(start) === DOLLAR) return false;

  const end = findClosing(src, start, posMax, marker);
  if (end === -1 || end === start) return false;
  // `$x $` closes nothing, and `$5 and $10` is two prices.
  if (isSpace(src.charCodeAt(end - 1))) return false;
  if (isDigit(src.charCodeAt(end + marker.length))) return false;

  if (!silent) {
    const token = state.push(double ? "math_inline_double" : "math_inline", "math", 0);
    token.markup = marker;
    token.content = src.slice(start, end);
  }
  state.pos = end + marker.length;
  return true;
}

/** A `$$ … $$` display block, on one line or many. */
function blockRule(
  state: StateBlock,
  startLine: number,
  endLine: number,
  silent: boolean,
): boolean {
  // Four spaces of indent is an indented code block, whatever it contains.
  if ((state.sCount[startLine] ?? 0) - state.blkIndent >= 4) return false;

  const begin = (state.bMarks[startLine] ?? 0) + (state.tShift[startLine] ?? 0);
  const max = state.eMarks[startLine] ?? begin;
  if (begin + 2 > max) return false;
  if (state.src.charCodeAt(begin) !== DOLLAR || state.src.charCodeAt(begin + 1) !== DOLLAR) {
    return false;
  }
  // Silent mode only asks "does a block start here", which decides whether this
  // can interrupt a paragraph. It can.
  if (silent) return true;

  let firstLine = state.src.slice(begin + 2, max);
  let lastLine = "";
  let found = false;
  let line = startLine;

  if (firstLine.trim().endsWith("$$")) {
    firstLine = firstLine.trim().slice(0, -2);
    found = true;
  }

  while (!found) {
    line += 1;
    if (line >= endLine) break;
    const from = (state.bMarks[line] ?? 0) + (state.tShift[line] ?? 0);
    const to = state.eMarks[line] ?? from;
    // A line that dedents out of the current container ends the block.
    if (from < to && (state.tShift[line] ?? 0) < state.blkIndent) break;

    const text = state.src.slice(from, to);
    if (text.trim().endsWith("$$")) {
      lastLine = state.src.slice(from, state.src.lastIndexOf("$$", to));
      found = true;
    }
  }

  // `keepLastLF: false`, so the parts join cleanly and `$$\nx\n$$` and `$$x$$`
  // yield the same TeX rather than differing by a trailing newline.
  const middle = state.getLines(startLine + 1, line, state.blkIndent, false);
  const parts = [firstLine, middle, lastLine].filter((part) => part.trim() !== "");

  const token = state.push("math_block", "math", 0);
  token.block = true;
  token.markup = "$$";
  token.content = parts.join("\n");
  token.map = [startLine, found ? line + 1 : line];
  state.line = found ? line + 1 : line;
  return true;
}

/** Options for {@link mathMarkdownIt}. */
export interface MathMarkdownItOptions {
  /** Which delimiters to recognise. Only `"dollars"` today. */
  readonly delimiters?: "dollars" | undefined;
}

/**
 * The markdown-it half of the math plugin.
 *
 * ```ts
 * const doc = parseMarkdown(source, { plugins: [mathMarkdownIt()] });
 * ```
 *
 * Produces tokens only. Filling each equation's `omml` slot is a separate,
 * asynchronous pass — see {@link import("./document.js").convertDocumentMath} —
 * because the TeX engine arrives through a dynamic `import()` and markdown-it
 * rules cannot await.
 *
 * @param options - See {@link MathMarkdownItOptions}.
 * @returns A plugin for `ParseOptions.plugins` or `md.use`.
 */
export function mathMarkdownIt(_options: MathMarkdownItOptions = {}): MarkdownItPlugin {
  return (md: MarkdownIt): void => {
    // Before `escape`, so the rule sees the `$` before `\$` handling does;
    // `$` is already one of markdown-it's text terminators, so the text rule
    // stops for us.
    md.inline.ruler.before("escape", "math_inline", inlineRule);
    md.block.ruler.before("fence", "math_block", blockRule, {
      alt: ["paragraph", "reference", "blockquote", "list"],
    });

    // downword never renders HTML, but a host may hand this plugin to its own
    // markdown-it. Without these the default renderer would emit a bare
    // `<math>` tag and drop the TeX.
    md.renderer.rules["math_inline"] = (tokens, index) =>
      `<span class="math math-inline">${md.utils.escapeHtml(tokens[index]?.content ?? "")}</span>`;
    md.renderer.rules["math_inline_double"] = (tokens, index) =>
      `<span class="math math-display">${md.utils.escapeHtml(tokens[index]?.content ?? "")}</span>`;
    md.renderer.rules["math_block"] = (tokens, index) =>
      `<div class="math math-display">${md.utils.escapeHtml(tokens[index]?.content ?? "")}</div>\n`;
  };
}
