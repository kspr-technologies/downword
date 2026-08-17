import type { Token } from "markdown-it";

import { createSlugger, type Slugger } from "../model.js";
import type { ResolvedParseOptions } from "./options.js";
import {
  parseWarningSeverity,
  type ParseWarningCode,
  type ParseWarningHandler,
} from "./warnings.js";

/**
 * Per-document parser state: the heading slugger, the warning sink and the few
 * options the walkers need.
 *
 * One instance per `parseMarkdown` call. Nothing here is shared between
 * documents, which is what makes heading anchors deterministic (`setup`,
 * `setup-1`, …) no matter how many documents a process converts.
 */
export class ParseContext {
  /** Allocates document-unique heading anchors. */
  readonly slugger: Slugger = createSlugger();

  /**
   * Whether a newline inside a paragraph becomes a `hardBreak` rather than a
   * `softBreak`.
   *
   * markdown-it's own `breaks` option is a **renderer** setting: it still emits
   * `softbreak` tokens and only changes how its HTML renderer prints them. This
   * parser never runs that renderer, so the option has to be applied here or it
   * would silently do nothing.
   */
  readonly hardBreaks: boolean;

  private readonly onWarning: ParseWarningHandler | null;

  /**
   * 1-based source line of the block currently being read.
   *
   * markdown-it only attaches `token.map` to *block* tokens, so an inline
   * warning would otherwise have no position at all. Remembering the enclosing
   * block's line gives every warning something useful to point at.
   */
  private blockLine: number | null = null;

  constructor(options: ResolvedParseOptions) {
    this.hardBreaks = options.breaks;
    this.onWarning = options.onWarning;
  }

  /** Records `token`'s source line as the fallback position for warnings. */
  noteLine(token: Token): void {
    if (token.map !== null) this.blockLine = token.map[0] + 1;
  }

  /**
   * Reports a non-fatal problem.
   *
   * @param code - Machine-readable kind.
   * @param message - Human-readable description.
   * @param token - Token the problem is about; its own line wins when it has one.
   */
  warn(code: ParseWarningCode, message: string, token: Token | null = null): void {
    if (this.onWarning === null) return;

    const line = token !== null && token.map !== null ? token.map[0] + 1 : this.blockLine;
    this.onWarning({ code, severity: parseWarningSeverity(code), message, line });
  }
}
