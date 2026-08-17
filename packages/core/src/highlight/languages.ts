/**
 * The lazy language registry.
 *
 * ## Why a hand-written table of `() => import(...)` thunks
 *
 * `highlight.js` ships two entry points: `highlight.js` (every grammar, ~1 MB
 * minified) and `highlight.js/lib/core` (the parser alone, ~35 kB) plus one
 * module per grammar. Only the second is viable here — a converter that pays a
 * megabyte so that one fence in ten can be coloured is a bad trade in a browser
 * tab, and `downword`'s bundle budget forbids it outright.
 *
 * That leaves *how* to reach a grammar module. Three options, and the tradeoff
 * is real:
 *
 * 1. **`import(`.../languages/${id}`)`** — one line, no table. But the
 *    specifier is not statically analysable: webpack turns it into a context
 *    module that eagerly bundles all 192 grammars, Vite refuses to transform it
 *    without a glob hint, and Rollup emits a warning and leaves it as a runtime
 *    import that will 404 in a bundled app. Rejected.
 * 2. **A table of literal `() => import("…/typescript")` thunks** — every
 *    specifier is a literal, so every bundler splits each grammar into its own
 *    chunk and *loads none of them* until a matching fence appears. The cost is
 *    this file: 62 lines of table, plus 62 chunks in the build output that most
 *    apps will never fetch. That is the trade taken here.
 * 3. **Host-supplied registry** — pass {@link HighlighterOptions.languages} to
 *    replace or extend the table. An app that only ever renders TypeScript can
 *    ship exactly one grammar chunk; an app that already bundles highlight.js
 *    can hand its own modules over and this table stays untouched.
 *
 * ## Which 62
 *
 * The full set is 192 grammars, most of which (`gams`, `mizar`, `rib`) will
 * never appear in a document a person pastes out of an LLM chat. The list below
 * is the intersection of "commonly written in markdown" and "highlight.js
 * supports it": the mainstream programming languages, the shells, the config
 * and data formats, and the markup languages. Anything outside it renders as
 * plain monospace, which is a correct rendering — never an error.
 */

import type { LanguageModule } from "./engine.js";

/** Loads one grammar module. Must resolve to highlight.js's default export. */
export type LanguageLoader = () => Promise<LanguageModule>;

/**
 * Canonical language name -> lazy grammar loader.
 *
 * Keys are highlight.js's own registration names; aliases live in
 * {@link LANGUAGE_ALIASES} so that one grammar is never registered twice.
 */
export const LANGUAGE_LOADERS: Readonly<Record<string, LanguageLoader>> = {
  apache: () => import("highlight.js/lib/languages/apache"),
  bash: () => import("highlight.js/lib/languages/bash"),
  c: () => import("highlight.js/lib/languages/c"),
  clojure: () => import("highlight.js/lib/languages/clojure"),
  cmake: () => import("highlight.js/lib/languages/cmake"),
  coffeescript: () => import("highlight.js/lib/languages/coffeescript"),
  cpp: () => import("highlight.js/lib/languages/cpp"),
  csharp: () => import("highlight.js/lib/languages/csharp"),
  css: () => import("highlight.js/lib/languages/css"),
  dart: () => import("highlight.js/lib/languages/dart"),
  diff: () => import("highlight.js/lib/languages/diff"),
  dockerfile: () => import("highlight.js/lib/languages/dockerfile"),
  dos: () => import("highlight.js/lib/languages/dos"),
  elixir: () => import("highlight.js/lib/languages/elixir"),
  erlang: () => import("highlight.js/lib/languages/erlang"),
  fsharp: () => import("highlight.js/lib/languages/fsharp"),
  go: () => import("highlight.js/lib/languages/go"),
  gradle: () => import("highlight.js/lib/languages/gradle"),
  graphql: () => import("highlight.js/lib/languages/graphql"),
  groovy: () => import("highlight.js/lib/languages/groovy"),
  haskell: () => import("highlight.js/lib/languages/haskell"),
  http: () => import("highlight.js/lib/languages/http"),
  ini: () => import("highlight.js/lib/languages/ini"),
  java: () => import("highlight.js/lib/languages/java"),
  javascript: () => import("highlight.js/lib/languages/javascript"),
  json: () => import("highlight.js/lib/languages/json"),
  julia: () => import("highlight.js/lib/languages/julia"),
  kotlin: () => import("highlight.js/lib/languages/kotlin"),
  latex: () => import("highlight.js/lib/languages/latex"),
  less: () => import("highlight.js/lib/languages/less"),
  lisp: () => import("highlight.js/lib/languages/lisp"),
  lua: () => import("highlight.js/lib/languages/lua"),
  makefile: () => import("highlight.js/lib/languages/makefile"),
  markdown: () => import("highlight.js/lib/languages/markdown"),
  matlab: () => import("highlight.js/lib/languages/matlab"),
  nginx: () => import("highlight.js/lib/languages/nginx"),
  nix: () => import("highlight.js/lib/languages/nix"),
  objectivec: () => import("highlight.js/lib/languages/objectivec"),
  ocaml: () => import("highlight.js/lib/languages/ocaml"),
  perl: () => import("highlight.js/lib/languages/perl"),
  pgsql: () => import("highlight.js/lib/languages/pgsql"),
  php: () => import("highlight.js/lib/languages/php"),
  powershell: () => import("highlight.js/lib/languages/powershell"),
  properties: () => import("highlight.js/lib/languages/properties"),
  protobuf: () => import("highlight.js/lib/languages/protobuf"),
  python: () => import("highlight.js/lib/languages/python"),
  r: () => import("highlight.js/lib/languages/r"),
  ruby: () => import("highlight.js/lib/languages/ruby"),
  rust: () => import("highlight.js/lib/languages/rust"),
  scala: () => import("highlight.js/lib/languages/scala"),
  scheme: () => import("highlight.js/lib/languages/scheme"),
  scss: () => import("highlight.js/lib/languages/scss"),
  shell: () => import("highlight.js/lib/languages/shell"),
  sql: () => import("highlight.js/lib/languages/sql"),
  swift: () => import("highlight.js/lib/languages/swift"),
  typescript: () => import("highlight.js/lib/languages/typescript"),
  vbnet: () => import("highlight.js/lib/languages/vbnet"),
  verilog: () => import("highlight.js/lib/languages/verilog"),
  vim: () => import("highlight.js/lib/languages/vim"),
  wasm: () => import("highlight.js/lib/languages/wasm"),
  xml: () => import("highlight.js/lib/languages/xml"),
  yaml: () => import("highlight.js/lib/languages/yaml"),
};

/**
 * Alias -> canonical name, for the languages in {@link LANGUAGE_LOADERS}.
 *
 * highlight.js registers a grammar's own `aliases` automatically, but only
 * *after* it is loaded — and knowing which module to load for `ts` is exactly
 * the question that has to be answered first. So the aliases are mirrored here.
 *
 * Machine-extracted from `highlight.js@11.11.1` (each grammar's `aliases`
 * array, verified collision-free across the 62 languages above) and then
 * extended by hand with the handful of spellings people write that
 * highlight.js does not itself claim.
 */
export const LANGUAGE_ALIASES: Readonly<Record<string, string>> = {
  apacheconf: "apache",
  atom: "xml",
  bat: "dos",
  "c#": "csharp",
  "c++": "cpp",
  cc: "cpp",
  cjs: "javascript",
  clj: "clojure",
  "cmake.in": "cmake",
  cmd: "dos",
  coffee: "coffeescript",
  console: "shell",
  cs: "csharp",
  cson: "coffeescript",
  cts: "typescript",
  cxx: "cpp",
  docker: "dockerfile",
  edn: "clojure",
  erl: "erlang",
  ex: "elixir",
  exs: "elixir",
  "f#": "fsharp",
  fs: "fsharp",
  gemspec: "ruby",
  golang: "go",
  gql: "graphql",
  gyp: "python",
  h: "c",
  "h++": "cpp",
  hh: "cpp",
  hpp: "cpp",
  hs: "haskell",
  html: "xml",
  https: "http",
  hxx: "cpp",
  iced: "coffeescript",
  ipython: "python",
  irb: "ruby",
  js: "javascript",
  jsonc: "json",
  jsp: "java",
  jsx: "javascript",
  kt: "kotlin",
  kts: "kotlin",
  mak: "makefile",
  make: "makefile",
  md: "markdown",
  mjs: "javascript",
  mk: "makefile",
  mkd: "markdown",
  mkdown: "markdown",
  ml: "ocaml",
  mm: "objectivec",
  mts: "typescript",
  nginxconf: "nginx",
  nixos: "nix",
  "obj-c": "objectivec",
  "obj-c++": "objectivec",
  objc: "objectivec",
  "objective-c++": "objectivec",
  patch: "diff",
  pl: "perl",
  plist: "xml",
  pluto: "lua",
  pm: "perl",
  podspec: "ruby",
  postgres: "pgsql",
  postgresql: "pgsql",
  proto: "protobuf",
  ps: "powershell",
  ps1: "powershell",
  pwsh: "powershell",
  py: "python",
  rb: "ruby",
  rs: "rust",
  rss: "xml",
  scm: "scheme",
  sh: "bash",
  shellsession: "shell",
  sv: "verilog",
  svg: "xml",
  svh: "verilog",
  tex: "latex",
  thor: "ruby",
  toml: "ini",
  ts: "typescript",
  tsx: "typescript",
  v: "verilog",
  vb: "vbnet",
  wsf: "xml",
  xhtml: "xml",
  xjb: "xml",
  xsd: "xml",
  xsl: "xml",
  yml: "yaml",
  zsh: "bash",

  /* Spellings highlight.js does not register, but people write anyway. */
  "docker-compose": "yaml",
  dotenv: "properties",
  env: "properties",
  htm: "xml",
  "objective-c": "objectivec",
  "shell-session": "shell",
  terminal: "shell",
  vue: "xml",
};

/**
 * Ids that mean "do not highlight this".
 *
 * Distinct from an *unknown* id: these are an explicit request for plain
 * monospace, so they short-circuit before the engine is even loaded and never
 * raise a `language-unknown` warning. `mermaid` is here because downword
 * renders those fences with its own plugin, not as code.
 */
export const PLAIN_LANGUAGES: ReadonlySet<string> = new Set([
  "plaintext",
  "plain",
  "text",
  "txt",
  "none",
  "nohighlight",
  "no-highlight",
  "raw",
  "output",
  "mermaid",
]);

/**
 * Grammars that must be registered alongside another one to colour the code
 * *embedded* in it.
 *
 * Several highlight.js grammars hand a region off to a second grammar —
 * `<script>` inside HTML, the command after a `$ ` prompt in a shell session,
 * a `RUN` line in a Dockerfile. That hand-off is resolved at parse time by
 * `getLanguage(name)`, so if the embedded grammar is not registered the region
 * silently renders as plain text; there is no error and no warning.
 *
 * This table is deliberately *short*. Thirteen of the 62 grammars declare a
 * `subLanguage` somewhere, but most of those are speculative — `typescript`
 * declares `xml`, `css` and `graphql` for JSX and tagged template literals, and
 * paying three extra chunk loads on every single TypeScript fence to colour a
 * construct that appears in a minority of them is a bad trade. Listed here are
 * only the pairs where the embedded language is the *point* of the outer one.
 */
export const LANGUAGE_DEPENDENCIES: Readonly<Record<string, readonly string[]>> = {
  /** `<script>` and `<style>` are most of what makes an HTML fence interesting. */
  xml: ["javascript", "css"],
  /** A shell *session* is a prompt plus a bash command; without bash it is prose. */
  shell: ["bash"],
  /** Every `RUN` line is bash. */
  dockerfile: ["bash"],
  /** Markdown embeds raw HTML by design. */
  markdown: ["xml"],
  /** CoffeeScript's whole reason for existing. */
  coffeescript: ["javascript"],
};

/**
 * Grammars auto-detection considers when it is switched on.
 *
 * Auto-detection has to *run* every candidate grammar over the source and
 * compare relevance scores, so the candidate set is both a CPU cost and — since
 * each one is a lazily loaded chunk — a network cost. Eleven of the most common
 * languages keeps both bounded while covering the overwhelming majority of
 * unlabelled fences. Override with `autoDetect: [...]`.
 */
export const AUTO_DETECT_LANGUAGES: readonly string[] = [
  "typescript",
  "javascript",
  "python",
  "bash",
  "json",
  "yaml",
  "xml",
  "sql",
  "java",
  "go",
  "css",
];

/** What {@link resolveLanguageId} decided about an id. */
export type LanguageResolution =
  /** An explicit "leave it plain" id such as `text`. */
  | { readonly kind: "plain" }
  /** A supported grammar, with the loader that fetches it. */
  | { readonly kind: "language"; readonly name: string; readonly load: LanguageLoader }
  /** Not in the registry. The caller falls back to plain and warns. */
  | { readonly kind: "unknown" };

/**
 * Maps a normalised id onto the registry.
 *
 * Alias resolution runs *before* the loader lookup and also *after* it fails,
 * so a host-supplied registry can register a grammar under an alias
 * (`languages: { ts: … }`) and still be found.
 */
export function resolveLanguageId(
  id: string,
  loaders: Readonly<Record<string, LanguageLoader>>,
  aliases: Readonly<Record<string, string>>,
): LanguageResolution {
  if (PLAIN_LANGUAGES.has(id)) return { kind: "plain" };

  const direct = loaders[id];
  if (direct !== undefined) return { kind: "language", name: id, load: direct };

  const canonical = aliases[id];
  if (canonical !== undefined) {
    const load = loaders[canonical];
    if (load !== undefined) return { kind: "language", name: canonical, load };
  }

  return { kind: "unknown" };
}
