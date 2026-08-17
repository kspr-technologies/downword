/**
 * Reading highlight.js's token stream instead of its HTML.
 *
 * highlight.js's advertised output is a string of `<span class="hljs-keyword">`
 * markup, and the obvious adapter regexes that apart. Three reasons not to:
 *
 *  - **It is lossy by construction.** `toHTML()` HTML-escapes the source, so
 *    `a < b && c` comes back as `a &lt; b &amp;&amp; c` and every span's text
 *    has to be un-escaped again. Entity handling is exactly the kind of detail
 *    that produces a `&amp;` in someone's Word document six months later.
 *  - **Scopes nest.** `title.function` inside `params` inside `function` is
 *    nested `<span>`s; flattening that with a regex means writing a parser
 *    badly, and writing it well means writing a real tokenizer for output that
 *    was a tree five microseconds ago.
 *  - **It is unnecessary.** `hljs.configure({ __emitter })` is a documented,
 *    supported option whose entire purpose is to let a caller receive the token
 *    stream directly. The tree never becomes HTML at all.
 *
 * So: no HTML is produced, none is parsed, and no escaping round-trip happens.
 * {@link SpanEmitter} records `(scope, text)` pairs as the parser walks, and
 * `toHTML()` returns `""` because nothing reads it.
 */

import type { HighlightEmitter } from "./engine.js";

/** One contiguous run of source text under a single innermost scope. */
export interface RawSpan {
  /** highlight.js's scope name (`hljs-` already stripped), or `null`. */
  readonly scope: string | null;
  /** Literal source text. Never entity-encoded. */
  readonly text: string;
}

/**
 * A highlight.js emitter that collects flat `(scope, text)` spans.
 *
 * Flat, not a tree: the renderer needs one `<w:r>` per contiguous run of
 * identical formatting, so nesting is collapsed to the *innermost* scope as the
 * parser walks. That matches how highlight.js's own CSS resolves — the deepest
 * `<span>` wins — and it means the output needs no post-order traversal.
 *
 * The parser calls the constructor itself (`new options.__emitter(options)`),
 * including once per embedded sub-language, which is why the options argument
 * is accepted and ignored rather than being replaced with a factory.
 */
export class SpanEmitter implements HighlightEmitter {
  /** Spans in source order. Concatenating `.text` reproduces the input. */
  readonly spans: RawSpan[] = [];

  private readonly stack: string[] = [];

  constructor(_options: unknown) {}

  private get scope(): string | null {
    return this.stack.length === 0 ? null : (this.stack[this.stack.length - 1] ?? null);
  }

  addText(text: string): void {
    if (text === "") return;
    this.spans.push({ scope: this.scope, text });
  }

  openNode(scope: string): void {
    this.stack.push(scope);
  }

  closeNode(): void {
    this.stack.pop();
  }

  startScope(scope: string): void {
    this.openNode(scope);
  }

  endScope(): void {
    this.closeNode();
  }

  /**
   * Splices an embedded language's spans in (`<script>` in HTML, `#{}` in Ruby,
   * `$(…)` in shell).
   *
   * The sub-parse ran with its own emitter and its own scope stack, so its
   * spans arrive already scoped and are taken verbatim. Text the sub-language
   * left unscoped inherits *this* emitter's current scope, which is what the
   * nested-`<span>` CSS would have done.
   *
   * A foreign emitter (possible only if a host reconfigures `__emitter`
   * mid-flight) contributes nothing rather than throwing: losing the colour on
   * an embedded fragment beats failing the document.
   */
  __addSublanguage(emitter: HighlightEmitter, _name: string | undefined): void {
    if (!(emitter instanceof SpanEmitter)) return;
    const outer = this.scope;
    for (const span of emitter.spans) {
      this.spans.push(span.scope === null ? { scope: outer, text: span.text } : span);
    }
  }

  finalize(): void {
    // Nothing to close: the parser balances every openNode with a closeNode,
    // and an unbalanced stack only affects scoping, never the text.
  }

  /**
   * Always `""`.
   *
   * highlight.js assigns this to `result.value`, which this adapter never
   * reads. Returning the empty string skips building and escaping a second
   * full copy of the source — a measurable saving on a large fence, and the
   * reason the emitter swap is a performance win as well as a correctness one.
   */
  toHTML(): string {
    return "";
  }
}
