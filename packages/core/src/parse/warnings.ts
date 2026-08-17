/**
 * Non-fatal problems the parser reports while it walks the token stream.
 *
 * The model deliberately carries no source positions (see `src/model.ts`), so
 * anything worth telling the user about has to be reported *while the parser
 * still holds the token* — which is what {@link ParseWarning.line} is for.
 * Parsing never throws on bad input: markdown has no syntax errors, only
 * constructs this converter chooses not to represent.
 */

import type { BaseWarning, WarningSeverity } from "../warnings.js";

/**
 * Machine-readable warning kind.
 *
 * Kept to a small closed set so callers can filter or suppress by code instead
 * of matching on human-readable prose.
 */
export type ParseWarningCode =
  /**
   * Raw HTML was encountered and turned into an `htmlBlock`/`htmlInline` node.
   * Emitted only when `html: true`, or when an injected plugin synthesises HTML
   * tokens. The *renderer* decides whether to drop it; the parser only reports.
   */
  | "raw-html"
  /**
   * A markdown-it token type this parser does not know about — almost always an
   * injected plugin whose tokens have no model counterpart. The token (and any
   * balanced region it opens) is skipped.
   */
  | "unsupported-token"
  /**
   * A token appeared somewhere the grammar says it cannot, e.g. an unbalanced
   * `strong_close` or a `td_open` outside a row. Indicates a plugin that
   * rewrote the stream badly, and is skipped defensively.
   */
  | "unexpected-token"
  /**
   * The source contained characters XML 1.0 cannot represent — a C0 control
   * other than tab/LF/CR, an unpaired surrogate, or `U+FFFE`/`U+FFFF`. Each was
   * replaced with `U+FFFD`, because writing one verbatim produces a `.docx`
   * Word refuses to open at all. See `src/xml-text.ts`.
   */
  | "unrepresentable-character"
  /**
   * A `[^x]: …` definition that nothing in the document references.
   *
   * `markdown-it-footnote` only emits the definitions its `footnote_tail` rule
   * finds in `env.footnotes.list`, and a label enters that list the first time
   * a `[^x]` *reference* is resolved against it. A definition nobody cites is
   * therefore filtered out of the token stream before any token for it exists —
   * its text is gone, and the parser is the only place that can still say so.
   */
  | "footnote-unreferenced"
  /**
   * Blocks nested deeper than markdown-it's ceiling (`MAX_BLOCK_NESTING`, 100
   * containers). markdown-it stops tokenizing at that depth and **discards the
   * rest of the enclosing region** — `">".repeat(120) + " text"` parses to an
   * empty document — so this is real content loss, not a reshaping. The
   * message names how many source lines went and where they started.
   */
  | "nesting-limit";

/**
 * Which parse warnings mean tokens were skipped rather than merely reshaped.
 *
 * Total by construction; see `src/warnings.ts` for where the line is drawn.
 */
const PARSE_WARNING_SEVERITY: Readonly<Record<ParseWarningCode, WarningSeverity>> = {
  // The token (and any region it opened) never reached the model.
  "unsupported-token": "error",
  "unexpected-token": "error",
  // markdown-it threw the source lines away before a token ever existed.
  "nesting-limit": "error",
  "footnote-unreferenced": "error",
  // The text is all present; only its representation changed.
  "raw-html": "notice",
  "unrepresentable-character": "notice",
};

/** Every parse warning code. See `RENDER_WARNING_CODES` for why this exists. */
export const PARSE_WARNING_CODES: readonly ParseWarningCode[] = Object.freeze(
  Object.keys(PARSE_WARNING_SEVERITY) as ParseWarningCode[],
);

/** The severity of a parse warning code. Total by construction. */
export function parseWarningSeverity(code: ParseWarningCode): WarningSeverity {
  return PARSE_WARNING_SEVERITY[code];
}

/** One non-fatal problem found while parsing. See {@link BaseWarning}. */
export interface ParseWarning extends BaseWarning {
  /** Machine-readable kind; see {@link ParseWarningCode}. */
  readonly code: ParseWarningCode;
  /**
   * 1-based line in the *source markdown*, or `null` when markdown-it attached
   * no source map to the token (true of most inline tokens, in which case the
   * enclosing block's line is used instead).
   */
  readonly line: number | null;
}

/** Callback handed a {@link ParseWarning} as soon as it is discovered. */
export type ParseWarningHandler = (warning: ParseWarning) => void;
