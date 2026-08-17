/**
 * Making arbitrary text safe to put inside an OOXML part.
 *
 * A `.docx` is a zip of XML 1.0 documents, and XML 1.0 cannot represent every
 * string JavaScript can hold. Its `Char` production is
 *
 * ```text
 * Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
 * ```
 *
 * so a C0 control other than tab/LF/CR, an unpaired surrogate, or either of the
 * two noncharacters `U+FFFE`/`U+FFFF` has **no escape** — not even a numeric
 * character reference. Writing one produces a part that every conforming parser
 * rejects, which in practice means Word refuses to open the file at all
 * ("we're sorry, we can't open … because we found a problem with its
 * contents"), losing the whole document rather than one character.
 *
 * Markdown pasted out of a chat window, a terminal capture or a PDF extractor
 * routinely contains these: `U+0001`, `U+000B` and `U+000C` all survive a
 * copy-paste, and a string sliced through a surrogate pair by a naive
 * truncation leaves a lone high surrogate behind.
 *
 * So every such character is replaced with `U+FFFD REPLACEMENT CHARACTER`, the
 * substitution Unicode defines for exactly this situation and the one
 * CommonMark already mandates for `U+0000`. The replacement is **one UTF-16
 * unit for one UTF-16 unit**, so source offsets — and therefore markdown-it's
 * line numbers and every warning that quotes one — are unchanged.
 */

/** What an unrepresentable character becomes: `U+FFFD REPLACEMENT CHARACTER`. */
export const XML_REPLACEMENT_CHARACTER = "�";

/**
 * Whether a *non-surrogate* UTF-16 unit is allowed in XML 1.0 character data.
 *
 * Surrogates are the caller's business: paired they encode a perfectly legal
 * astral character, unpaired they are illegal, and only the caller knows which.
 */
function isXmlChar(unit: number): boolean {
  if (unit === 0x09 || unit === 0x0a || unit === 0x0d) return true;
  if (unit < 0x20) return false;
  if (unit <= 0xd7ff) return true;
  // 0xD800-0xDFFF never reaches here; 0xFFFE and 0xFFFF are excluded.
  return unit >= 0xe000 && unit <= 0xfffd;
}

/** How many UTF-16 units the character starting at `index` occupies, and whether it is legal. */
function measure(value: string, index: number): { readonly width: number; readonly ok: boolean } {
  const unit = value.charCodeAt(index);
  if (unit >= 0xd800 && unit <= 0xdbff) {
    const low = index + 1 < value.length ? value.charCodeAt(index + 1) : 0;
    const paired = low >= 0xdc00 && low <= 0xdfff;
    // A well-formed pair is always a legal XML character (U+10000-U+10FFFF).
    return paired ? { width: 2, ok: true } : { width: 1, ok: false };
  }
  if (unit >= 0xdc00 && unit <= 0xdfff) return { width: 1, ok: false }; // orphan low surrogate
  return { width: 1, ok: isXmlChar(unit) };
}

/**
 * Replaces every character XML 1.0 cannot carry with {@link XML_REPLACEMENT_CHARACTER}.
 *
 * Returns the original string when there is nothing to do, so the common case
 * costs one scan and no allocation.
 *
 * @param value - Any string.
 * @returns A string of the same length that is safe to write into an OOXML part.
 */
export function sanitizeXmlText(value: string): string {
  let out: string | null = null;
  let copied = 0;
  let index = 0;

  while (index < value.length) {
    const { width, ok } = measure(value, index);
    if (!ok) {
      out = (out ?? "") + value.slice(copied, index) + XML_REPLACEMENT_CHARACTER;
      copied = index + width;
    }
    index += width;
  }

  return out === null ? value : out + value.slice(copied);
}

/**
 * Counts the characters {@link sanitizeXmlText} would replace.
 *
 * Separate from the substitution itself because the number is only ever needed
 * to phrase a warning, and paying for it on clean input would be wasteful.
 *
 * @param value - Any string.
 * @returns How many characters cannot be represented in XML 1.0.
 */
export function countXmlUnrepresentable(value: string): number {
  let count = 0;
  let index = 0;

  while (index < value.length) {
    const { width, ok } = measure(value, index);
    if (!ok) count += 1;
    index += width;
  }

  return count;
}
