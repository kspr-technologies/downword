/**
 * The two engines, loaded lazily and wrapped in the smallest interface that
 * does the job.
 *
 * ```text
 *   TeX ──temml──▶ MathML ──mathml2omml──▶ <m:oMath> ──▶ word/document.xml
 * ```
 *
 * ## Both are optional peer dependencies
 *
 * `temml` (~250 kB) and `mathml2omml` (~60 kB) are reached by bare specifier
 * from a `import()` inside this file, and from nowhere else in the package.
 * That is what keeps them out of `downword`'s main entry — a guarantee
 * `tests/bundle.test.ts` enforces by bundling `src/index.ts` and failing if
 * either name appears in the import graph — and it means an application that
 * never writes an equation never downloads a TeX parser.
 *
 * A host that already bundles them, or a browser with no bundler, can hand the
 * modules over directly with {@link MathEngineOptions.load} instead.
 *
 * ## `temml`, not KaTeX
 *
 * KaTeX renders to HTML plus a stylesheet; there is no way back from that to
 * OOXML. `temml` renders the same TeX to **MathML**, which is a semantic tree
 * that maps onto OMML element for element — and it does it without a DOM, so
 * the same code runs in a worker and on a server.
 *
 * ## Everything about the call is defensive
 *
 * `mathml2omml` writes `console.warn("Type not supported: …")` straight to the
 * host's console for every MathML element it has no OMML for. Since it is
 * synchronous, and JavaScript is not re-entrant, swapping `console.warn` for
 * the duration of the call is safe and turns that noise into a diagnostic on
 * the package's own warning channel. See {@link captureConsoleWarnings}.
 */

/** How `temml` should render one equation. */
export interface TexRenderOptions {
  /** `$$…$$` renders in display style: bigger operators, limits above and below. */
  readonly displayMode: boolean;
  /** `\newcommand`-style macros, TeX bodies only. A fresh copy per equation. */
  readonly macros: Record<string, string>;
}

/**
 * The two conversions, and nothing else.
 *
 * Both are synchronous and both may throw — a `ParseError` from `temml`, or a
 * `RangeError` from either when an equation nests deeply enough to exhaust the
 * stack. Callers treat any throw as "this equation stays literal text".
 */
export interface MathEngine {
  /** TeX -> MathML. Throws `temml`'s `ParseError` on TeX it cannot read. */
  renderToMathml(tex: string, options: TexRenderOptions): string;
  /** MathML -> an `<m:oMath>` fragment, with entity decoding disabled. */
  mathmlToOmml(mathml: string): string;
}

/** Supplies a {@link MathEngine}. Called at most once per converter. */
export type MathEngineLoader = () => Promise<MathEngine>;

/* -------------------------------------------------------------------------- */
/* Console capture                                                             */
/* -------------------------------------------------------------------------- */

/**
 * Runs `body` with `console.warn` diverted into an array.
 *
 * Only sound because `mathml2omml` is synchronous: nothing else can run between
 * the swap and the restore, so no unrelated log can be swallowed and no
 * concurrent conversion can see the replacement. `console.error` and
 * `console.log` are left alone — the converter does not use them.
 */
export function captureConsoleWarnings<T>(body: () => T): {
  readonly result: T;
  readonly messages: readonly string[];
} {
  const messages: string[] = [];
  const target: Partial<Console> | undefined = typeof console === "undefined" ? undefined : console;

  if (target === undefined || typeof target.warn !== "function") {
    return { result: body(), messages };
  }

  const original = target.warn.bind(console);
  target.warn = (...args: readonly unknown[]): void => {
    messages.push(args.map((arg) => (typeof arg === "string" ? arg : String(arg))).join(" "));
  };
  try {
    return { result: body(), messages };
  } finally {
    target.warn = original;
  }
}

/* -------------------------------------------------------------------------- */
/* The default loader                                                          */
/* -------------------------------------------------------------------------- */

/**
 * `import("temml")` + `import("mathml2omml")`, wrapped.
 *
 * Rejects with whatever the runtime throws when a peer dependency is missing;
 * the caller turns that into an `engine-unavailable` warning and leaves every
 * equation as literal TeX.
 */
export const loadMathEngine: MathEngineLoader = async (): Promise<MathEngine> => {
  const [temml, mathml2omml] = await Promise.all([import("temml"), import("mathml2omml")]);
  // temml publishes named exports in its .d.ts but only a default at runtime.
  const renderToString = temml.default.renderToString;
  const { mml2omml } = mathml2omml;

  return {
    renderToMathml(tex, options) {
      return renderToString(tex, {
        // Self-closing tags and no HTML-only entities: the output has to parse
        // as XML, not merely as HTML.
        xml: true,
        displayMode: options.displayMode,
        // Errors belong on the warning channel, not rendered in red inside the
        // document.
        throwOnError: true,
        // An <annotation> holding the original TeX would be walked into the
        // OMML as a second copy of the source text.
        annotate: false,
        // Keeps \includegraphics and \href refused. The default, made explicit
        // because a document is not a trusted input.
        trust: false,
        strict: false,
        // Line-break hints have no OMML equivalent.
        wrap: "none",
        macros: options.macros,
      });
    },

    mathmlToOmml(mathml) {
      // disableDecode is load-bearing, not a tweak: without it the converter
      // decodes temml's escaped MathML and writes the result raw, which turns
      // `$<w:p>…</w:p>$` into live markup in word/document.xml. See `omml.ts`.
      return mml2omml(mathml, { disableDecode: true });
    },
  };
};
