/**
 * The default syntax highlighter: highlight.js, loaded lazily, theme-driven.
 *
 * ```ts
 * import { renderDocument, prepareHighlights } from "downword/render";
 * import { createHighlighter } from "downword/highlight";
 *
 * const highlights = await prepareHighlights(doc, createHighlighter());
 * const file = renderDocument(doc, { highlights });
 * ```
 *
 * ## Who chooses the colours
 *
 * This adapter reports *what each token is* — `keyword`, `string`,
 * `title.function` — and the renderer's {@link
 * import("../render/theme.js").Theme.codePalette} decides what colour that is.
 * One palette therefore covers both of the renderer's highlighting paths, and
 * `convert(md, { theme })` governs a highlighted fence just as it governs
 * everything else. What this module keeps for itself is the part a palette
 * cannot express: **bold keywords and italic comments**, so a greyscale
 * printout still separates them. See {@link HighlighterOptions.scopeStyles} to
 * take the colours back.
 *
 * ## Everything about this module is lazy
 *
 * Nothing in `src/highlight/**` imports highlight.js statically. The engine
 * arrives through `await import("highlight.js/lib/core")` on the first fence
 * that needs it, and each grammar through its own `await import(...)` on the
 * first fence written in that language. Importing this module therefore costs
 * a few kilobytes of tables, and a document with no code blocks — or only
 * unlabelled ones — never fetches highlight.js at all.
 *
 * That is not a micro-optimisation: `downword`'s bundle budget is enforced in
 * CI against the package entry point, and a static `import "highlight.js"` is
 * roughly a megabyte. highlight.js is an **optional peer dependency**; an app
 * that does not install it still builds, and every code block renders as plain
 * monospace.
 *
 * ## Node and the browser
 *
 * The engine is reached by bare specifier, which resolves natively in Node and
 * through any bundler in the browser. Two environments need help:
 *
 *  - a browser with **no bundler** — add an import map, or pass
 *    {@link HighlighterOptions.load}; and
 *  - an app that **already has** highlight.js loaded — pass `load` so the two
 *    copies do not both ship.
 *
 * Nothing here touches `fs`, `document` or `window`.
 *
 * ## It never throws
 *
 * A missing engine, an unknown language, a grammar that fails to fetch, an
 * absurdly large fence, an internal highlight.js error: every one of them
 * degrades to plain monospace and (where it is actionable) an
 * {@link HighlightWarning}. `prepareHighlights` already isolates a throwing
 * highlighter per-block, but a converter should not need that safety net.
 */

import type { Highlighter, HighlightSpan } from "../render/types.js";
import type {
  HighlightEngine,
  HighlightEngineLoader,
  LanguageDefinition,
  LanguageModule,
} from "./engine.js";
import { SpanEmitter } from "./emitter.js";
import { normalizeLanguageId } from "./info.js";
import {
  AUTO_DETECT_LANGUAGES,
  LANGUAGE_ALIASES,
  LANGUAGE_DEPENDENCIES,
  LANGUAGE_LOADERS,
  resolveLanguageId,
  type LanguageLoader,
  type LanguageResolution,
} from "./languages.js";
import { PRINT_SCOPE_STYLES, type ScopeStyle } from "./palette.js";
import { plainSpans, toHighlightSpans } from "./spans.js";

export { SpanEmitter, type RawSpan } from "./emitter.js";
export { normalizeLanguageId } from "./info.js";
export {
  AUTO_DETECT_LANGUAGES,
  LANGUAGE_ALIASES,
  LANGUAGE_DEPENDENCIES,
  LANGUAGE_LOADERS,
  PLAIN_LANGUAGES,
  resolveLanguageId,
  type LanguageLoader,
  type LanguageResolution,
} from "./languages.js";
export {
  PRINT_CODE_PALETTE,
  PRINT_INK,
  PRINT_SCOPE_STYLES,
  resolveScopeStyle,
  type PrintInkRole,
  type ScopeStyle,
} from "./palette.js";
export { plainSpans, toHighlightSpans } from "./spans.js";
export type {
  HighlightEmitter,
  HighlightEngine,
  HighlightEngineLoader,
  HighlightEngineResult,
  LanguageDefinition,
  LanguageModule,
} from "./engine.js";

/* -------------------------------------------------------------------------- */
/* Warnings                                                                    */
/* -------------------------------------------------------------------------- */

/** Machine-readable reason a {@link HighlightWarning} was raised. */
export type HighlightWarningCode =
  /** `import("highlight.js/lib/core")` failed. Raised once per highlighter. */
  | "engine-unavailable"
  /** The fence named a language that is not in the registry. */
  | "language-unknown"
  /** The grammar is in the registry but its module failed to load. */
  | "language-load-failed"
  /** The fence exceeded {@link HighlighterOptions.maxLength}. */
  | "code-too-large"
  /** highlight.js threw, or returned a token stream this adapter cannot read. */
  | "highlight-failed";

/** A non-fatal problem while highlighting. The block renders plain monospace. */
export interface HighlightWarning {
  readonly code: HighlightWarningCode;
  readonly message: string;
  /** The normalised language id, or `null` for an unlabelled fence. */
  readonly lang: string | null;
}

/* -------------------------------------------------------------------------- */
/* Options                                                                     */
/* -------------------------------------------------------------------------- */

/** Options for {@link createHighlighter}. */
export interface HighlighterOptions {
  /**
   * Scope -> run formatting, **and the switch that takes colour away from the
   * renderer's theme**.
   *
   * Left unset (the default), spans carry their scope name and the renderer
   * colours them from
   * {@link import("../render/theme.js").Theme.codePalette} — so
   * `convert(md, { theme: "print", highlighter: createHighlighter() })` really
   * does print in the print inks, and `theme: "default"` really does use the
   * screen ones. {@link PRINT_SCOPE_STYLES} still supplies the weight and slant
   * (bold keywords, italic comments), because those are not colour and a
   * `.docx` is a print artefact under either palette.
   *
   * Set it, and you own the palette: spans carry the colours you specify, no
   * scope is reported, and the theme is not consulted for code at all. Passing
   * `PRINT_SCOPE_STYLES` explicitly therefore pins the print inks regardless of
   * theme.
   *
   * **Replaces** the default map rather than merging with it, so a partial
   * object yields a mostly-uncoloured listing. To adjust a few scopes, spread:
   * `{ ...PRINT_SCOPE_STYLES, comment: { color: "808080" } }`.
   */
  readonly scopeStyles?: Readonly<Record<string, ScopeStyle>> | undefined;
  /**
   * Canonical name -> grammar loader. Defaults to {@link LANGUAGE_LOADERS}.
   *
   * **Replaces** the default table, which is the point: passing
   * `{ typescript: () => import("highlight.js/lib/languages/typescript") }`
   * means the build contains exactly one grammar chunk instead of 62. Spread
   * {@link LANGUAGE_LOADERS} to add to the defaults instead of narrowing them.
   */
  readonly languages?: Readonly<Record<string, LanguageLoader>> | undefined;
  /** Alias -> canonical name. Defaults to {@link LANGUAGE_ALIASES}; replaces it. */
  readonly aliases?: Readonly<Record<string, string>> | undefined;
  /**
   * Grammars to register alongside another one so that *embedded* code is
   * coloured too. Defaults to {@link LANGUAGE_DEPENDENCIES}; replaces it.
   *
   * Set to `{}` to never load more than one grammar per fence, at the cost of
   * plain `<script>` bodies inside an HTML fence and plain commands inside a
   * shell session.
   */
  readonly dependencies?: Readonly<Record<string, readonly string[]>> | undefined;
  /**
   * Guess the language of a fence that names none (or names an unknown one).
   *
   * **Off by default**, deliberately. Auto-detection runs *every* candidate
   * grammar over the source and keeps the highest-scoring one, so switching it
   * on turns one lazy chunk into eleven and one parse into eleven — and it is
   * wrong often enough on short snippets (a three-line fence of `key: value`
   * scores as half a dozen languages) that mislabelled colour is a real cost.
   * The fence's info string is authoritative when it exists.
   *
   * `true` uses {@link AUTO_DETECT_LANGUAGES}; an array names the candidates.
   */
  readonly autoDetect?: boolean | readonly string[] | undefined;
  /**
   * Longest fence to highlight, in UTF-16 code units. Defaults to 2 000 000
   * (roughly 25 000-50 000 lines of code).
   *
   * A backstop, not a normal limit: measured against highlight.js 11.11.1, a
   * 50 000-line TypeScript fence (4.4 MB) parses in ~1.4 s and yields ~1.15 M
   * tokens, so the default leaves plenty of headroom while still refusing to
   * spend unbounded time and memory on a pathological paste. Above it the block
   * renders plain and a `code-too-large` warning is raised.
   */
  readonly maxLength?: number | undefined;
  /**
   * Supplies the highlight.js core. Defaults to
   * `() => import("highlight.js/lib/core")`.
   *
   * The returned engine is never used directly — `newInstance()` is called on
   * it so that this adapter's emitter configuration cannot leak into a copy of
   * highlight.js the host is also using.
   */
  readonly load?: HighlightEngineLoader | undefined;
  /** Called once per non-fatal problem. */
  readonly onWarning?: ((warning: HighlightWarning) => void) | undefined;
}

/* -------------------------------------------------------------------------- */
/* Engine plumbing                                                             */
/* -------------------------------------------------------------------------- */

/** A highlight.js instance plus the bookkeeping for lazily registered grammars. */
interface Engine {
  readonly hljs: HighlightEngine;
  readonly registered: Set<string>;
  /** In-flight (and settled) registrations, so a grammar loads exactly once. */
  readonly loading: Map<string, Promise<boolean>>;
}

const defaultLoad: HighlightEngineLoader = async () =>
  (await import("highlight.js/lib/core")).default;

/** Normalises the three module shapes a grammar can arrive in. */
function unwrapLanguage(module: LanguageModule): LanguageDefinition | null {
  if (typeof module === "function") return module;
  const inner = module.default;
  if (typeof inner === "function") return inner;
  return typeof inner.default === "function" ? inner.default : null;
}

/* -------------------------------------------------------------------------- */
/* The adapter                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Creates a {@link Highlighter} backed by highlight.js.
 *
 * The result is stateful — it caches the engine and every grammar it has
 * loaded — so create one and reuse it across documents. Creating several is
 * harmless but each keeps its own registry, re-running the (cheap) grammar
 * registration against a module the runtime has already cached.
 *
 * Pair it with `prepareHighlights(doc, highlighter)`: `highlight()` is async
 * (it may have to fetch a grammar) and `renderDocument` is synchronous, so
 * passing this straight to `options.highlighter` would raise a
 * `highlighter-async` warning and colour nothing.
 */
export function createHighlighter(options: HighlighterOptions = {}): Highlighter {
  const styles = options.scopeStyles ?? PRINT_SCOPE_STYLES;
  // Scopes travel with the spans unless the caller took the palette over; see
  // HighlighterOptions.scopeStyles for why that is the switch.
  const reportScopes = options.scopeStyles === undefined;
  const loaders = options.languages ?? LANGUAGE_LOADERS;
  const aliases = options.aliases ?? LANGUAGE_ALIASES;
  const dependencies = options.dependencies ?? LANGUAGE_DEPENDENCIES;
  const maxLength = options.maxLength ?? 2_000_000;
  const load = options.load ?? defaultLoad;
  const autoDetect =
    options.autoDetect === true
      ? AUTO_DETECT_LANGUAGES
      : options.autoDetect === false || options.autoDetect === undefined
        ? []
        : options.autoDetect;

  const warn = (code: HighlightWarningCode, message: string, lang: string | null): void => {
    options.onWarning?.({ code, message, lang });
  };

  let engineOnce: Promise<Engine | null> | null = null;

  const engine = async (): Promise<Engine | null> => {
    engineOnce ??= (async () => {
      try {
        const hljs = (await load()).newInstance();
        // The whole reason this adapter can read tokens instead of HTML.
        hljs.configure({ __emitter: SpanEmitter });
        return {
          hljs,
          registered: new Set<string>(),
          loading: new Map<string, Promise<boolean>>(),
        };
      } catch (error) {
        warn(
          "engine-unavailable",
          `highlight.js could not be loaded, so code blocks will render unhighlighted: ${String(error)}`,
          null,
        );
        return null;
      }
    })();
    return engineOnce;
  };

  const register = async (
    active: Engine,
    name: string,
    loader: LanguageLoader,
  ): Promise<boolean> => {
    if (active.registered.has(name)) return true;

    let inFlight = active.loading.get(name);
    if (inFlight === undefined) {
      inFlight = (async () => {
        try {
          const definition = unwrapLanguage(await loader());
          if (definition === null) return false;
          active.hljs.registerLanguage(name, definition);
          active.registered.add(name);
          return true;
        } catch {
          return false;
        }
      })();
      active.loading.set(name, inFlight);
    }
    return inFlight;
  };

  /**
   * Registers a grammar plus whatever it embeds.
   *
   * Dependencies are best-effort: a `<script>` block losing its colour because
   * the JavaScript chunk 404'd is not worth a warning, let alone a failure of
   * the outer grammar.
   */
  const registerWithDependencies = async (
    active: Engine,
    resolution: Extract<LanguageResolution, { kind: "language" }>,
  ): Promise<boolean> => {
    if (!(await register(active, resolution.name, resolution.load))) return false;

    for (const dependency of dependencies[resolution.name] ?? []) {
      const resolved = resolveLanguageId(dependency, loaders, aliases);
      if (resolved.kind === "language") await register(active, resolved.name, resolved.load);
    }
    return true;
  };

  /** Runs one parse and converts its tokens, or returns `null` to fall back. */
  const parse = (
    code: string,
    lang: string | null,
    run: () => { readonly _emitter: unknown },
  ): readonly HighlightSpan[] | null => {
    let emitter: unknown;
    try {
      emitter = run()._emitter;
    } catch (error) {
      warn("highlight-failed", `highlight.js failed on this block: ${String(error)}`, lang);
      return null;
    }
    if (!(emitter instanceof SpanEmitter)) {
      warn("highlight-failed", "highlight.js did not use the configured emitter", lang);
      return null;
    }
    return toHighlightSpans(emitter.spans, code, styles, reportScopes);
  };

  const detect = async (
    active: Engine,
    code: string,
    lang: string | null,
  ): Promise<readonly HighlightSpan[] | null> => {
    const available: string[] = [];
    for (const name of autoDetect) {
      const resolved = resolveLanguageId(name, loaders, aliases);
      if (resolved.kind !== "language") continue;
      if (await registerWithDependencies(active, resolved)) available.push(resolved.name);
    }
    if (available.length === 0) return null;
    return parse(code, lang, () => active.hljs.highlightAuto(code, available));
  };

  return {
    async highlight(code: string, lang: string | null): Promise<readonly HighlightSpan[]> {
      const fallback = plainSpans(code);
      if (code === "") return fallback;

      const id = normalizeLanguageId(lang);

      if (code.length > maxLength) {
        warn(
          "code-too-large",
          `code block of ${code.length} characters exceeds maxLength ${maxLength}; rendering it unhighlighted`,
          id,
        );
        return fallback;
      }

      const resolution: LanguageResolution =
        id === null ? { kind: "unknown" } : resolveLanguageId(id, loaders, aliases);

      // An explicit `text`/`none` fence is a request, not a failure: no engine
      // load, no warning.
      if (resolution.kind === "plain") return fallback;

      // With auto-detection off there is nothing an engine could do for a fence
      // whose language is absent or unsupported - so do not fetch one. This is
      // what keeps a document full of unlabelled fences at zero network cost.
      if (resolution.kind === "unknown" && autoDetect.length === 0) {
        // An unlabelled fence is ordinary markdown; only a *named* language
        // that is not supported is worth telling the caller about.
        if (id !== null) {
          warn(
            "language-unknown",
            `no grammar registered for "${id}"; rendering this block unhighlighted`,
            id,
          );
        }
        return fallback;
      }

      const active = await engine();
      if (active === null) return fallback;

      if (resolution.kind === "language") {
        if (await registerWithDependencies(active, resolution)) {
          return (
            parse(code, id, () =>
              active.hljs.highlight(code, { language: resolution.name, ignoreIllegals: true }),
            ) ?? fallback
          );
        }
        warn(
          "language-load-failed",
          `the "${resolution.name}" grammar could not be loaded; rendering this block unhighlighted`,
          id,
        );
        return fallback;
      }

      const detected = await detect(active, code, id);
      if (detected !== null) return detected;

      if (id !== null) {
        warn(
          "language-unknown",
          `no grammar registered for "${id}" and auto-detection found nothing; rendering this block unhighlighted`,
          id,
        );
      }
      return fallback;
    },
  };
}
