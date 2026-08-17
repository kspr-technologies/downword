/**
 * The slice of the highlight.js runtime this adapter actually touches.
 *
 * highlight.js ships a `.d.ts`, but its `./lib/core` subpath carries no `types`
 * condition and no ambient declaration, so `import("highlight.js/lib/core")`
 * has nothing to resolve. Rather than reach for `any` (banned) or drag the
 * whole `HLJSApi` surface in, the handful of members used here are declared
 * structurally and bound to the real module in `vendor.d.ts`. (The grammar
 * subpath needs no such help: highlight.js declares
 * `highlight.js/lib/languages/*` as a wildcard ambient module itself.)
 *
 * A second, larger reason to own these types: they are the seam that lets a
 * host supply *its own* highlight.js. A browser app that already loads
 * highlight.js from a CDN, or a bundler that cannot see through a bare
 * specifier, can pass {@link HighlightEngineLoader} instead and never touch
 * this package's dynamic imports.
 */

/**
 * highlight.js's emitter protocol, as the engine actually calls it.
 *
 * The published `Emitter` interface lists only six methods, but the parser also
 * calls `openNode`/`closeNode` directly (`core.js` `startNewMode`,
 * `processContinuations`, `endOfMode`). An emitter that implements only the
 * documented six throws at runtime, so both pairs are declared here.
 */
export interface HighlightEmitter {
  /** Appends literal source text at the current scope. */
  addText(text: string): void;
  /** Pushes a scope (`"keyword"`, `"title.function"`, …) onto the stack. */
  openNode(scope: string): void;
  /** Pops the innermost scope. */
  closeNode(): void;
  /** `openNode` under a different name; used for keyword matches. */
  startScope(scope: string): void;
  /** `closeNode` under a different name. */
  endScope(): void;
  /** Splices an embedded language's emitter (e.g. `<script>` inside HTML). */
  __addSublanguage(emitter: HighlightEmitter, name: string | undefined): void;
  /** Called once, after the last token. */
  finalize(): void;
  /** The engine assigns this to `result.value`; unused by this adapter. */
  toHTML(): string;
}

/** How highlight.js instantiates the configured emitter: `new opts.__emitter(opts)`. */
export interface HighlightEmitterConstructor {
  new (options: unknown): HighlightEmitter;
}

/**
 * A highlight.js language definition — the default export of every
 * `highlight.js/lib/languages/*` module.
 *
 * Deliberately opaque: the adapter only ever hands one straight back to
 * `registerLanguage`, and modelling `Language`/`Mode` would couple this file to
 * highlight.js's internals for no gain.
 *
 * The parameter is `never`, not `unknown`, and that is load-bearing. Under
 * `strictFunctionTypes` a parameter is contravariant, so highlight.js's own
 * `LanguageFn = (hljs: HLJSApi) => Language` is assignable to
 * `(hljs: never) => unknown` but *not* to `(hljs: unknown) => unknown`. `never`
 * is also the honest signature: nothing in this package ever calls one.
 */
export type LanguageDefinition = (hljs: never) => unknown;

/**
 * What a grammar module resolves to.
 *
 * Three shapes rather than one because a grammar can arrive through three
 * different module systems: highlight.js's `es/languages/*.js` are real ESM
 * (`{ default: fn }`), its `lib/languages/*.js` are CommonJS
 * (`module.exports = fn`, so a bundler's interop may hand over the bare
 * function), and a double-wrapped `{ default: { default: fn } }` is what some
 * CJS-to-ESM interop layers produce. Accepting all three is four lines in
 * `unwrapLanguage` and removes an entire category of "works in vitest, breaks
 * in the CJS build" bug.
 */
export type LanguageModule =
  | LanguageDefinition
  | { readonly default: LanguageDefinition }
  | { readonly default: { readonly default: LanguageDefinition } };

/** The subset of highlight.js's `HighlightResult` this adapter reads. */
export interface HighlightEngineResult {
  /** The language actually used; `undefined` when auto-detection found none. */
  readonly language?: string | undefined;
  /** Auto-detection score. Only meaningful for `highlightAuto`. */
  readonly relevance: number;
  /** True when the source violated the grammar and parsing stopped early. */
  readonly illegal: boolean;
  /**
   * The emitter instance that received the token stream.
   *
   * Underscored because highlight.js considers it private, but it is the only
   * way to read structured tokens instead of an HTML string, and the
   * `__emitter` option exists precisely so callers can substitute their own.
   */
  readonly _emitter: HighlightEmitter;
}

/** The subset of highlight.js's public API this adapter calls. */
export interface HighlightEngine {
  /** A fresh registry + options object, isolated from the module singleton. */
  newInstance(): HighlightEngine;
  configure(options: { readonly __emitter: HighlightEmitterConstructor }): void;
  registerLanguage(name: string, language: LanguageDefinition): void;
  /** Non-`undefined` when `name` is a registered language *or one of its aliases*. */
  getLanguage(name: string): unknown;
  listLanguages(): readonly string[];
  highlight(
    code: string,
    options: { readonly language: string; readonly ignoreIllegals?: boolean },
  ): HighlightEngineResult;
  highlightAuto(code: string, languageSubset?: readonly string[]): HighlightEngineResult;
  readonly versionString: string;
}

/**
 * Supplies the highlight.js runtime.
 *
 * The default implementation is `() => import("highlight.js/lib/core")`, kept
 * dynamic so highlight.js never enters this package's static import graph.
 */
export type HighlightEngineLoader = () => Promise<HighlightEngine>;
