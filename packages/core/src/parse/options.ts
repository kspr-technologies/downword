import MarkdownItConstructor, { type MarkdownIt } from "markdown-it";
import footnotePlugin from "markdown-it-footnote";

import { metadata, type DocumentMetadata } from "../model.js";
import { parseWarningSeverity, type ParseWarningHandler } from "./warnings.js";

/**
 * A markdown-it plugin, already bound to whatever options it needs.
 *
 * Plugins that take options are injected as a closure rather than as a
 * `[plugin, options]` tuple, which keeps this type free of `any`:
 *
 * ```ts
 * import katex from "@vscode/markdown-it-katex";
 *
 * parseMarkdown(source, {
 *   plugins: [(md) => md.use(katex.default ?? katex, { throwOnError: false })],
 * });
 * ```
 */
export type MarkdownItPlugin = (md: MarkdownIt) => void;

/** Options accepted by `parseMarkdown`. */
export interface ParseOptions {
  /**
   * Let raw HTML through as `HtmlBlockNode` / `HtmlInlineNode` nodes.
   *
   * Defaults to `false`, which is markdown-it's safe default: HTML is then left
   * as literal text, exactly as GitHub renders it in a sanitised context. When
   * enabled the parser only *models* the markup and emits a `raw-html` warning
   * — deciding whether to render or drop it is the renderer's policy, kept in
   * one place instead of being silently made here.
   */
  readonly html?: boolean | undefined;
  /**
   * Turn bare URLs and e-mail addresses into links. Defaults to `true`, because
   * LLM output is full of unlinked URLs.
   */
  readonly linkify?: boolean | undefined;
  /**
   * Treat every newline inside a paragraph as a hard break. Defaults to
   * `false` (CommonMark), which models single newlines as `SoftBreakNode`s and
   * leaves the join-or-preserve decision to the renderer.
   *
   * Note that markdown-it's own `breaks` option is a *renderer* setting — it
   * still emits `softbreak` tokens — so this parser applies the flag itself
   * when mapping them; see `ParseContext.hardBreaks`.
   */
  readonly breaks?: boolean | undefined;
  /**
   * Smart quotes, en/em dashes and ellipses. Defaults to `false` so the output
   * contains exactly the characters the author typed.
   */
  readonly typographer?: boolean | undefined;
  /**
   * Parse `[^1]` footnotes via `markdown-it-footnote`. Defaults to `true`.
   * Disabling it leaves footnote syntax as literal text.
   */
  readonly footnotes?: boolean | undefined;
  /**
   * Extra markdown-it plugins, applied in order after the built-ins.
   *
   * This is how math reaches the parser: `@vscode/markdown-it-katex` (or any
   * plugin emitting `math_inline`/`math_block` tokens) is injected here rather
   * than being a hard dependency of this package, and the resulting tokens are
   * mapped to `MathInlineNode` /
   * `MathBlockNode`.
   */
  readonly plugins?: readonly MarkdownItPlugin[] | undefined;
  /**
   * Document metadata (normally lifted from frontmatter by the caller). The
   * parser does not read frontmatter itself — that is a host concern, and YAML
   * is not a dependency of this package.
   */
  readonly metadata?: DocumentMetadata | undefined;
  /** Called for each non-fatal problem. See {@link ParseWarningHandler}. */
  readonly onWarning?: ParseWarningHandler | undefined;
}

/** {@link ParseOptions} with every default filled in. Total by construction. */
export interface ResolvedParseOptions {
  readonly html: boolean;
  readonly linkify: boolean;
  readonly breaks: boolean;
  readonly typographer: boolean;
  readonly footnotes: boolean;
  readonly plugins: readonly MarkdownItPlugin[];
  readonly metadata: DocumentMetadata;
  readonly onWarning: ParseWarningHandler | null;
}

/** Shared empty plugin list. */
const NO_PLUGINS: readonly MarkdownItPlugin[] = Object.freeze([]);

/**
 * How many nested block containers markdown-it will tokenize.
 *
 * markdown-it's own `"default"` preset happens to use 100 as well, but it is
 * pinned here rather than inherited, for the same reason `table` and
 * `strikethrough` are enabled explicitly: the grammar this parser accepts
 * should be stated in the source. It is also the number the nesting guard
 * below compares against, and the two must agree.
 *
 * A hundred containers is far past anything a human or an LLM writes — a
 * 50-deep bullet list is 100, since each level costs a list *and* an item —
 * but it is trivially reachable by a fuzzer or a paste gone wrong, and what
 * markdown-it does at the ceiling is throw the rest of the region away.
 */
export const MAX_BLOCK_NESTING = 100;

/**
 * Reports the source lines markdown-it discards when it hits
 * {@link MAX_BLOCK_NESTING}.
 *
 * `ParserBlock.tokenize` bails with `state.line = endLine` the moment
 * `state.level` reaches the ceiling — no token is produced and no error is
 * raised, so `">".repeat(120) + " text"` parses to a completely empty document
 * in silence. That is the one place in this pipeline where content can vanish
 * before a node ever exists, which is why it is caught by wrapping the method
 * rather than by inspecting the model afterwards: at this point the state still
 * knows exactly which lines are about to go.
 *
 * The wrapper is a single integer comparison on every other call, and is
 * installed only when someone is listening.
 */
function guardBlockNesting(md: MarkdownIt, onWarning: ParseWarningHandler): void {
  const tokenize = md.block.tokenize.bind(md.block);

  md.block.tokenize = (state, startLine, endLine) => {
    if (state.level >= MAX_BLOCK_NESTING) {
      // Mirror tokenize's own first-iteration test: it skips blank lines, then
      // stops without consuming anything if the region belongs to an outer
      // container. Only past both of those is anything actually thrown away.
      const first = state.skipEmptyLines(startLine);
      if (first < endLine && (state.sCount[first] ?? 0) >= state.blkIndent) {
        let dropped = 0;
        for (let line = first; line < endLine; line++) if (!state.isEmpty(line)) dropped++;
        onWarning({
          code: "nesting-limit",
          severity: parseWarningSeverity("nesting-limit"),
          message:
            `blocks nested deeper than ${MAX_BLOCK_NESTING} levels: markdown-it stopped ` +
            `parsing and discarded ${dropped} source line${dropped === 1 ? "" : "s"}`,
          line: first + 1,
        });
      }
    }
    tokenize(state, startLine, endLine);
  };
}

/**
 * Fills in every {@link ParseOptions} default.
 *
 * @param options - Any subset of the options.
 * @returns Total options; absent values become their default, never `undefined`.
 */
export function resolveParseOptions(options: ParseOptions = {}): ResolvedParseOptions {
  return {
    html: options.html ?? false,
    linkify: options.linkify ?? true,
    breaks: options.breaks ?? false,
    typographer: options.typographer ?? false,
    footnotes: options.footnotes ?? true,
    plugins: options.plugins ?? NO_PLUGINS,
    metadata: options.metadata ?? metadata(),
    onWarning: options.onWarning ?? null,
  };
}

/**
 * Builds the markdown-it instance the parser walks.
 *
 * markdown-it's `"default"` preset is CommonMark plus GFM tables and
 * strikethrough; both are re-enabled explicitly here so the grammar this parser
 * accepts is stated in the source rather than inherited from a preset that
 * could change.
 *
 * **Task lists are handled natively, not by `markdown-it-task-lists`.** That
 * plugin rewrites the list item's inline children into raw
 * `<input type="checkbox">` `html_inline` tokens and exposes no `checked` field
 * on any token, so a docx renderer would have to parse HTML back out of them.
 * It also keeps its configuration in module-level mutable state and misses
 * items with an empty body (`- [x]`). Detecting `[ ]`/`[x]` on the first text
 * token is both simpler and strictly more accurate — see `blocks.ts`. If the
 * plugin *is* injected through {@link ParseOptions.plugins} anyway, its
 * checkbox tokens are recognised too, so both paths produce the same model.
 *
 * When a warning handler is supplied it is also wired to the block nesting
 * ceiling; see {@link MAX_BLOCK_NESTING}. Callers who build their own instance
 * and hand the tokens to `parseTokens` do not get that check, because by then
 * the discarded lines are already gone.
 *
 * @param options - Resolved options; see {@link resolveParseOptions}.
 * @returns A configured, plugin-loaded markdown-it instance.
 */
export function createMarkdownIt(options: ResolvedParseOptions): MarkdownIt {
  const md = new MarkdownItConstructor({
    html: options.html,
    linkify: options.linkify,
    breaks: options.breaks,
    typographer: options.typographer,
    xhtmlOut: false,
  });

  // Pin the ceiling rather than inherit the preset's, so the guard below and
  // the parser cannot disagree about where it is.
  Object.assign(md.options, { maxNesting: MAX_BLOCK_NESTING });

  md.enable(["table", "strikethrough"]);

  if (options.footnotes) md.use(footnotePlugin);
  for (const plugin of options.plugins) md.use(plugin);

  // Last, so it wraps any tokenize a plugin installed rather than being
  // replaced by one.
  if (options.onWarning !== null) guardBlockNesting(md, options.onWarning);

  return md;
}
