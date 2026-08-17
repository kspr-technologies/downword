/**
 * The slice of the mermaid runtime this adapter actually touches, and the
 * default {@link MermaidRenderer} built on it.
 *
 * ### Why these types are owned here
 *
 * mermaid ships its own `.d.ts`, and importing it would drag mermaid's public
 * type surface — and, with `skipLibCheck` off in a consumer's project, its
 * dependency graph — into a package that must not depend on mermaid at all.
 * Two members are used (`initialize`, `render`), so two members are declared,
 * structurally, and bound to the real module in `vendor.d.ts` exactly as
 * `src/highlight/engine.ts` does for highlight.js.
 *
 * The larger reason is the same one as there: this is the seam that lets a host
 * supply *its own* mermaid. An app that already loads it from a CDN, a worker
 * that was handed the module, or a test that wants a two-line fake all pass a
 * {@link MermaidLoader} and this package's dynamic import is never reached.
 *
 * ### Why the import is dynamic
 *
 * mermaid is ~500 kB and an **optional** peer dependency. `import("mermaid")`
 * inside a function keeps it out of the static graph, so a bundler emits it as
 * its own chunk and a document with no diagrams never fetches it — the same
 * arrangement `downword/highlight` has with highlight.js, and one
 * `tests/bundle.test.ts` enforces for the core entry.
 */

import type { MermaidRenderRequest, MermaidRenderer, MermaidSvg } from "./types.js";

/* -------------------------------------------------------------------------- */
/* The runtime, structurally                                                   */
/* -------------------------------------------------------------------------- */

/**
 * mermaid's configuration object.
 *
 * Deliberately opaque: mermaid's config is a hundred keys deep, versioned, and
 * none of it means anything to downword — it is passed through untouched.
 * The two keys this module sets itself are documented on
 * {@link BrowserRendererOptions.config}.
 */
export type MermaidConfig = Readonly<Record<string, unknown>>;

/** What `mermaid.render()` resolves to. Only `svg` is read. */
export interface MermaidRenderOutput {
  readonly svg: string;
}

/** The subset of mermaid's public API this adapter calls. */
export interface MermaidApi {
  /** Must be called before the first `render`; `startOnLoad` is set to `false`. */
  initialize(config: MermaidConfig): void;
  /**
   * Renders one diagram.
   *
   * `id` becomes a real element id in the live document, and is written into
   * the SVG's own `id` and the CSS selectors inside it.
   */
  render(id: string, text: string): Promise<MermaidRenderOutput>;
}

/**
 * What the mermaid module resolves to.
 *
 * Two shapes rather than one: mermaid's ESM build is `{ default: api }`, and a
 * CJS-to-ESM interop layer (or a host handing over `window.mermaid`) can
 * produce the bare object. Accepting both is four lines and removes a whole
 * category of "works in the app, breaks in the bundle" bug.
 */
export type MermaidModule = MermaidApi | { readonly default: MermaidApi };

/**
 * Supplies the mermaid runtime.
 *
 * The default is `() => import("mermaid")`, kept dynamic so mermaid never
 * enters this package's static import graph.
 */
export type MermaidLoader = () => Promise<MermaidModule>;

function hasApiShape(value: unknown): value is MermaidApi {
  if (typeof value !== "object" || value === null) return false;
  const candidate = value as { readonly initialize?: unknown; readonly render?: unknown };
  return typeof candidate.initialize === "function" && typeof candidate.render === "function";
}

/** Unwraps `{ default: api }` / `api`, or throws with a message worth reading. */
export function unwrapMermaid(module: MermaidModule | unknown): MermaidApi {
  if (hasApiShape(module)) return module;
  if (typeof module === "object" && module !== null && "default" in module) {
    const inner: unknown = (module as { readonly default: unknown }).default;
    if (hasApiShape(inner)) return inner;
  }
  throw new Error(
    "the mermaid module does not look like mermaid: expected an object with initialize() and render()",
  );
}

/* -------------------------------------------------------------------------- */
/* The default renderer                                                        */
/* -------------------------------------------------------------------------- */

/** Options for {@link createBrowserMermaidRenderer}. */
export interface BrowserRendererOptions {
  /**
   * Extra mermaid configuration, merged over downword's own two defaults:
   *
   * - `startOnLoad: false` — downword calls `render` itself; letting mermaid
   *   scan the page would rewrite the host's DOM as a side effect of converting
   *   a document.
   * - `securityLevel: "strict"` — mermaid's own sanitiser, on. Diagram source
   *   arrives in pasted markdown, so it is untrusted input; `"loose"` lets a
   *   label carry HTML that is then inserted into the host page. Override it
   *   only for markdown you wrote.
   *
   * Everything else is passed straight through — `theme`, `flowchart`,
   * `fontFamily`, and so on.
   */
  readonly config?: MermaidConfig | undefined;
  /** Supplies the runtime. Defaults to `() => import("mermaid")`. See {@link MermaidLoader}. */
  readonly load?: MermaidLoader | undefined;
}

const DEFAULT_CONFIG: MermaidConfig = { startOnLoad: false, securityLevel: "strict" };

/**
 * `Error.name` used to mark "mermaid itself could not be loaded".
 *
 * The pass has to tell that apart from "this diagram does not parse": the first
 * is one fact about the whole run (report it once, stop trying), the second is
 * one broken diagram among many. A marker on `name` rather than a subclass,
 * because `instanceof` is not reliable across this package's dual ESM/CJS build
 * — the same reason `isDownwordError` is a structural check.
 */
export const MERMAID_ENGINE_ERROR = "MermaidEngineUnavailableError";

/** Wraps whatever the loader threw in a marked, readable error. */
export function engineUnavailableError(cause: unknown): Error {
  const detail = cause instanceof Error ? cause.message : String(cause);
  const error = new Error(
    `mermaid could not be loaded (${detail}). It is an optional peer dependency: ` +
      `install it with \`npm install mermaid\`, or pass your own renderer.`,
    { cause },
  );
  error.name = MERMAID_ENGINE_ERROR;
  return error;
}

/** Whether an error means "no mermaid runtime", rather than "no such diagram". */
export function isEngineUnavailable(error: unknown): boolean {
  return error instanceof Error && error.name === MERMAID_ENGINE_ERROR;
}

/** `() => import("mermaid")`, isolated so nothing else in the file mentions the specifier. */
const importMermaid: MermaidLoader = () => import("mermaid");

/**
 * Removes the element mermaid rendered into, if it survived.
 *
 * mermaid appends a temporary element with the caller's id to `document.body`
 * and normally removes it again — but not on every error path, and a leftover
 * makes the *next* render with the same id fail. Since the ids here are
 * deterministic (they have to be, or the SVG bytes would differ between two
 * conversions of the same document), cleaning up is not optional.
 */
function removeStrayElement(id: string): void {
  if (typeof document === "undefined") return;
  try {
    document.getElementById(id)?.remove();
    // mermaid 10/11 also parks a measuring container under this id.
    document.getElementById(`d${id}`)?.remove();
  } catch {
    /* a DOM that cannot be tidied is not this diagram's problem */
  }
}

/**
 * The default {@link MermaidRenderer}: real mermaid, real DOM.
 *
 * The runtime is loaded once, on the first diagram, and `initialize` is called
 * once with it. Diagrams are rendered one at a time by the pass, because
 * mermaid keeps module-level state and renders through the shared document.
 *
 * @param options - Configuration and the loader seam. See {@link BrowserRendererOptions}.
 * @returns A renderer whose `render` rejects with a readable error when mermaid
 *   is missing or the diagram does not parse.
 */
export function createBrowserMermaidRenderer(
  options: BrowserRendererOptions = {},
): MermaidRenderer {
  const load = options.load ?? importMermaid;
  const config: MermaidConfig = { ...DEFAULT_CONFIG, ...options.config };
  let engine: Promise<MermaidApi> | null = null;

  const ready = (): Promise<MermaidApi> => {
    if (engine === null) {
      engine = (async () => {
        const api = unwrapMermaid(await load());
        api.initialize(config);
        return api;
      })().catch((error: unknown) => {
        // Do not cache a failed load: a host that installs mermaid and retries
        // (or one whose CDN blipped) should get a second chance.
        engine = null;
        throw engineUnavailableError(error);
      });
    }
    return engine;
  };

  return {
    async render(request: MermaidRenderRequest): Promise<MermaidSvg> {
      const api = await ready();
      try {
        const output = await api.render(request.id, request.source);
        return { svg: output.svg, width: null, height: null };
      } finally {
        removeStrayElement(request.id);
      }
    },
  };
}
