/**
 * markdown -> intermediate model.
 *
 * The parser's whole job is to turn markdown-it's **flat** token array into the
 * nested tree in `src/model.ts`, and to do it without losing anything the
 * renderer will need. It never throws: markdown has no syntax errors, only
 * constructs this converter chooses not to represent, and those are reported
 * through {@link ParseOptions.onWarning} instead.
 *
 * ```ts
 * const document = parseMarkdown("# Hello\n\n- world\n");
 * ```
 *
 * Nothing here is re-exported from `src/index.ts` yet — the public API surface
 * is settled once the renderer lands, and the barrel is owned by that change.
 *
 * @module
 */

import type { Env, Token } from "markdown-it";

import { doc, type DocumentNode } from "../model.js";
import { countXmlUnrepresentable, sanitizeXmlText } from "../xml-text.js";
import { readBlocks } from "./blocks.js";
import { ParseContext } from "./context.js";
import { TokenCursor } from "./cursor.js";
import {
  createMarkdownIt,
  resolveParseOptions,
  type ParseOptions,
  type ResolvedParseOptions,
} from "./options.js";
import { parseWarningSeverity } from "./warnings.js";

export { createMarkdownIt, MAX_BLOCK_NESTING, resolveParseOptions } from "./options.js";
export type { MarkdownItPlugin, ParseOptions, ResolvedParseOptions } from "./options.js";
export { PARSE_WARNING_CODES, parseWarningSeverity } from "./warnings.js";
export type { ParseWarning, ParseWarningCode, ParseWarningHandler } from "./warnings.js";

/**
 * Parses markdown into a {@link DocumentNode}.
 *
 * Handles CommonMark plus GFM tables, strikethrough, task lists and autolinks;
 * footnotes via `markdown-it-footnote`; and math via any injected plugin that
 * emits `math_inline`/`math_block` tokens (see {@link ParseOptions.plugins}).
 *
 * Heading anchors are allocated with one slugger per document, so two headings
 * reading "Setup" become `setup` and `setup-1`, deterministically.
 *
 * The source is first swept for characters XML 1.0 cannot hold — C0 controls
 * other than tab/LF/CR, unpaired surrogates, `U+FFFE`/`U+FFFF` — which are
 * replaced with `U+FFFD` and reported once as `unrepresentable-character`. This
 * is CommonMark's own `U+0000` rule widened to everything OOXML shares the
 * problem with: emitting one verbatim yields a `.docx` Word refuses to open, so
 * the choice is one substituted character or no document at all. Because the
 * substitution is one UTF-16 unit for one, every source offset — and so every
 * warning's `line` — is unaffected.
 *
 * @param source - Raw markdown. May be empty; the result is then an empty document.
 * @param options - See {@link ParseOptions}. Every field has a default.
 * @returns The document. Footnote definitions stay in `children`, at the end,
 *   for the renderer to hoist.
 */
export function parseMarkdown(source: string, options: ParseOptions = {}): DocumentNode {
  const resolved = resolveParseOptions(options);
  const md = createMarkdownIt(resolved);
  const env: Env = {};

  const text = sanitizeXmlText(source);
  if (text !== source && resolved.onWarning !== null) {
    const count = countXmlUnrepresentable(source);
    resolved.onWarning({
      code: "unrepresentable-character",
      severity: parseWarningSeverity("unrepresentable-character"),
      message: `replaced ${count} character${count === 1 ? "" : "s"} that XML cannot represent with U+FFFD`,
      line: null,
    });
  }

  const tokens = md.parse(text, env);
  reportUnreferencedFootnotes(env, resolved);
  return parseTokens(tokens, resolved);
}

/**
 * Reports `[^x]: …` definitions that `markdown-it-footnote` threw away.
 *
 * A definition only reaches the token stream if some `[^x]` resolved against
 * it: `footnote_def` records the label as `env.footnotes.refs[":x"] = -1` and
 * parks its tokens, `footnote_ref` promotes it to an index into
 * `env.footnotes.list`, and `footnote_tail` re-emits only what is in that list.
 * A label still sitting at `-1` when parsing finishes is one whose whole body —
 * possibly several paragraphs of it — was filtered out of the stream, silently,
 * before the parser saw a single token. Nothing downstream can notice, because
 * from its point of view the definition was never written.
 *
 * The `env` shape is markdown-it's, not ours, so every field is checked at
 * runtime: an author who disables footnotes, or replaces the plugin, must get
 * silence rather than a crash.
 */
function reportUnreferencedFootnotes(env: Env, options: ResolvedParseOptions): void {
  const onWarning = options.onWarning;
  if (onWarning === null) return;

  const footnotes: unknown = (env as Record<string, unknown>)["footnotes"];
  if (typeof footnotes !== "object" || footnotes === null) return;

  const refs: unknown = (footnotes as Record<string, unknown>)["refs"];
  if (typeof refs !== "object" || refs === null) return;

  for (const [key, index] of Object.entries(refs as Record<string, unknown>)) {
    if (index !== -1) continue;
    // markdown-it-footnote prefixes the key with ':' to keep labels off
    // Object.prototype.
    const label = key.startsWith(":") ? key.slice(1) : key;
    onWarning({
      code: "footnote-unreferenced",
      severity: parseWarningSeverity("footnote-unreferenced"),
      message:
        `the footnote definition [^${label}] is never referenced, so markdown-it discarded it; ` +
        `add a [^${label}] marker in the text, or delete the definition`,
      line: null,
    });
  }
}

/**
 * Walks an already-produced markdown-it token stream into a
 * {@link DocumentNode}.
 *
 * For callers that own their own markdown-it instance (a host that shares one
 * across features, or a test that injects a synthetic stream). The instance
 * must have been built the way {@link createMarkdownIt} builds one, or tokens
 * this parser does not know about will simply be warned about and skipped.
 *
 * @param tokens - Block-level tokens from `md.parse()`.
 * @param options - Resolved options; see {@link resolveParseOptions}.
 */
export function parseTokens(tokens: readonly Token[], options: ResolvedParseOptions): DocumentNode {
  const context = new ParseContext(options);
  const cursor = new TokenCursor(tokens);

  return doc(readBlocks(cursor, context, null), options.metadata);
}
