/**
 * Turning scoped tokens into the renderer's {@link HighlightSpan}s.
 *
 * Two invariants matter more than anything else here:
 *
 *  1. **Concatenating every span's `text` must equal the input exactly.** The
 *     renderer splits the spans on `\n` to make one paragraph per line; a
 *     dropped character silently corrupts someone's code listing. Enforced by
 *     {@link reconcile} rather than assumed.
 *  2. **Adjacent spans with identical formatting must be merged.** Each span
 *     becomes a `<w:r>` in `document.xml`, and highlight.js happily emits five
 *     consecutive unscoped tokens for `; ` `\n` ` ` `= ` `…`. Merging is what
 *     keeps a 50 000-line fence from producing a million runs.
 *
 * "Identical formatting" is not always "identical colour": when the spans carry
 * their {@link HighlightSpan.scope} for the renderer's theme to colour, two
 * scopes that happen to share an ink *here* may not share one *there*. See
 * {@link createFormatResolver}.
 */

import type { HighlightSpan } from "../render/types.js";
import type { RawSpan } from "./emitter.js";
import { resolveScopeStyle, type ScopeStyle } from "./palette.js";

/** Everything about a span except its text: what one `<w:r>` will look like. */
type SpanFormat = Omit<HighlightSpan, "text">;

/** The "no formatting" span: inherit the CodeBlock style's colour and weight. */
const PLAIN_FORMAT: SpanFormat = {
  color: undefined,
  bold: undefined,
  italic: undefined,
  scope: undefined,
};

/** Identity of a span's formatting, for interning. */
function formatKey(format: SpanFormat): string {
  return `${format.color ?? ""}|${format.bold === true ? 1 : 0}|${
    format.italic === true ? 1 : 0
  }|${format.scope ?? ""}`;
}

const PLAIN_KEY = formatKey(PLAIN_FORMAT);

/**
 * Resolves scopes to span formatting, interning the results.
 *
 * Interning is what makes coalescing cheap *and* thorough: two adjacent tokens
 * that will render identically resolve to the identical object, so merging is
 * an `===`. What "identically" means depends on `reportScopes`:
 *
 *  - **off** — the styles carry the final colours, so `attr` and `number` (same
 *    ink) merge into one span. That is the shape a caller who supplied their
 *    own `scopeStyles` asked for.
 *  - **on** — the *theme* picks the colour later and may well give those two
 *    scopes different ones, so the scope is part of the identity and they stay
 *    apart. The renderer coalesces again once it knows the theme, so
 *    `document.xml` ends up with the same number of runs either way.
 */
function createFormatResolver(
  styles: Readonly<Record<string, ScopeStyle>>,
  reportScopes: boolean,
): (scope: string | null) => SpanFormat {
  const byScope = new Map<string, SpanFormat>();
  // Seeded so that a scope the palette does not name interns to the *same*
  // object an unscoped token gets, and the two merge. (Only reachable with
  // `reportScopes` off; with it on, an unnamed scope still has to stay
  // separate, because the theme may yet colour it.)
  const byValue = new Map<string, SpanFormat>([[PLAIN_KEY, PLAIN_FORMAT]]);

  return (scope) => {
    if (scope === null) return PLAIN_FORMAT;

    const cached = byScope.get(scope);
    if (cached !== undefined) return cached;

    const style: ScopeStyle = resolveScopeStyle(styles, scope) ?? {};
    const resolved: SpanFormat = {
      color: style.color,
      bold: style.bold,
      italic: style.italic,
      scope: reportScopes ? scope : undefined,
    };
    const key = formatKey(resolved);
    const interned = byValue.get(key) ?? resolved;
    byValue.set(key, interned);
    byScope.set(scope, interned);
    return interned;
  };
}

/**
 * Guarantees the spans reproduce `code`.
 *
 * highlight.js can return a *partial* token stream: when a grammar hits an
 * `illegal` match (or safe mode swallows an internal error) it returns the
 * emitter it had filled in so far. `ignoreIllegals: true` makes that rare, not
 * impossible, and the failure mode — a code block truncated mid-line — is far
 * worse than an uncoloured one.
 *
 * So the emitted text is checked against the source. A partial stream that is a
 * proper prefix keeps its colours and gets the remainder appended unstyled; a
 * stream that diverges some other way is discarded entirely.
 */
function reconcile(spans: readonly RawSpan[], code: string): readonly RawSpan[] {
  let emitted = 0;
  for (const span of spans) emitted += span.text.length;
  if (emitted === code.length) {
    // Same length is not the same string, but a highlighter that substitutes
    // characters without changing the length is a bug this cannot paper over -
    // and the full comparison below is only worth paying for when lengths differ.
    return spans;
  }
  if (emitted < code.length) {
    let joined = "";
    for (const span of spans) joined += span.text;
    if (code.startsWith(joined)) {
      return [...spans, { scope: null, text: code.slice(joined.length) }];
    }
  }
  return [{ scope: null, text: code }];
}

/**
 * Converts raw tokens into renderer spans: resolve, merge, verify.
 *
 * Every optional field is written explicitly as `undefined` when unset, which
 * `HighlightSpan` documents as the shape adapters compiled with
 * `exactOptionalPropertyTypes` should produce.
 *
 * @param spans - Raw `(scope, text)` tokens, in source order.
 * @param code - The source they came from; the result is checked against it.
 * @param styles - Scope -> run formatting. See `PRINT_SCOPE_STYLES`.
 * @param reportScopes - Carry {@link HighlightSpan.scope} through, so the
 *   renderer's theme can pick the colour. Defaults to `false`, which bakes in
 *   whatever `styles` says and leaves the theme out of it.
 */
export function toHighlightSpans(
  spans: readonly RawSpan[],
  code: string,
  styles: Readonly<Record<string, ScopeStyle>>,
  reportScopes = false,
): readonly HighlightSpan[] {
  const resolve = createFormatResolver(styles, reportScopes);
  const out: HighlightSpan[] = [];

  let pending: SpanFormat | null = null;
  let text = "";

  const flush = (): void => {
    if (pending === null) return;
    out.push({ text, ...pending });
    pending = null;
    text = "";
  };

  for (const span of reconcile(spans, code)) {
    const format = resolve(span.scope);
    if (format !== pending) {
      flush();
      pending = format;
    }
    text += span.text;
  }
  flush();

  return out;
}

/** The whole input as a single unformatted span. */
export function plainSpans(code: string): readonly HighlightSpan[] {
  return code === "" ? [] : [{ text: code, ...PLAIN_FORMAT }];
}
