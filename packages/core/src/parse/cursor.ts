import type { Token } from "markdown-it";

/**
 * A read-only, forward-only position in a markdown-it token stream.
 *
 * markdown-it emits a *flat* array in which nesting is encoded by
 * `token.nesting` (`1` open, `0` self-closing, `-1` close). Rebuilding the tree
 * therefore means a recursive descent over one shared cursor, and every reader
 * in `blocks.ts` follows the same contract:
 *
 * - it is called with the cursor sitting **on** its opening token;
 * - it consumes that token and everything up to and including its closer;
 * - it consumes **at least one** token, so no caller loop can spin forever.
 */
export class TokenCursor {
  private readonly tokens: readonly Token[];
  private position = 0;

  constructor(tokens: readonly Token[]) {
    this.tokens = tokens;
  }

  /** Index of the next token; only meaningful for progress assertions. */
  get index(): number {
    return this.position;
  }

  /** The next token without consuming it, or `null` at the end of the stream. */
  peek(): Token | null {
    return this.tokens[this.position] ?? null;
  }

  /** Consumes and returns the next token, or `null` at the end of the stream. */
  next(): Token | null {
    const token = this.peek();
    if (token !== null) this.position += 1;
    return token;
  }

  /**
   * Discards the next token — and, if it opens a region, everything through its
   * matching closer.
   *
   * Used for tokens the parser has no model for: skipping the whole balanced
   * region keeps the walker in sync instead of leaving orphaned close tokens to
   * be reported as further errors.
   */
  skip(): void {
    const token = this.next();
    if (token === null || token.nesting !== 1) return;

    let depth = 1;
    while (depth > 0) {
      const inner = this.next();
      if (inner === null) return;
      depth += inner.nesting;
    }
  }
}
