/**
 * Turning a fence info string into a language id.
 *
 * The model's `CodeBlockNode.lang` is already "first word of the info string",
 * but {@link import("../render/types.js").Highlighter} is a *public* interface:
 * anything may call `highlight(code, lang)`, including code that passes the
 * whole info string. Normalising defensively here costs one regex and removes a
 * whole class of "why is my ```ts title=foo block not coloured" reports.
 *
 * The syntaxes in the wild:
 *
 * | written                          | means                          | -> |
 * | -------------------------------- | ------------------------------ | -- |
 * | ` ```ts `                        | plain                          | `ts` |
 * | ` ```ts title=foo.ts `           | docusaurus / nextra            | `ts` |
 * | ` ```js{1,3-5} `                 | vuepress line highlighting     | `js` |
 * | ` ```{.python .numberLines} `    | pandoc attributes              | `python` |
 * | ` ```language-python `           | prism / rehype class name      | `python` |
 * | ` ```JSON `                      | any casing                     | `json` |
 * | ` ```c++ ` / ` ```f# `           | punctuation *inside* the name  | `c++` / `f#` |
 */

/**
 * Characters that may appear inside a language id.
 *
 * `+`, `#`, `-`, `.` and `_` are all load-bearing (`c++`, `f#`, `obj-c`,
 * `cmake.in`, `php_template`), which is why the id cannot simply be split on
 * "the first non-alphanumeric character". Everything else — whitespace, `{`,
 * `=`, `,`, `:`, `"` — terminates it.
 */
const ID_PATTERN = /^[a-z0-9+#._-]+/;

/** Class-name prefixes markdown pipelines add before the real language. */
const CLASS_PREFIXES = ["language-", "lang-", "highlight-source-"] as const;

/**
 * Extracts a lowercase language id from a fence info string.
 *
 * Returns `null` for an absent, blank or unusable info string — the caller
 * treats that the same as an unknown language: plain monospace, no throw.
 *
 * Pure and total; it never touches the language registry, so `"klingon"`
 * normalises to `"klingon"` rather than to `null`. Deciding whether an id is
 * *supported* is {@link import("./languages.js").resolveLanguageId}'s job.
 */
export function normalizeLanguageId(info: string | null | undefined): string | null {
  if (info === null || info === undefined) return null;

  // Pandoc writes `{.python .numberLines}`; strip the brace and the leading dot
  // so the ordinary "first token" rule applies to the rest.
  let rest = info.trim().toLowerCase();
  if (rest.startsWith("{")) rest = rest.slice(1).trimStart();
  while (rest.startsWith(".")) rest = rest.slice(1);

  const matched = ID_PATTERN.exec(rest)?.[0];
  if (matched === undefined) return null;

  let id = matched;
  for (const prefix of CLASS_PREFIXES) {
    if (id.length > prefix.length && id.startsWith(prefix)) {
      id = id.slice(prefix.length);
      break;
    }
  }

  // A trailing dot is punctuation, not part of the id: `ts.` -> `ts`. A leading
  // one was already consumed above.
  while (id.endsWith(".")) id = id.slice(0, -1);

  return id === "" ? null : id;
}
