/**
 * The renderer's last line of defence against a string XML cannot carry.
 *
 * `parseMarkdown` sanitizes every string it puts in the model (`parse/index.ts`)
 * and `resolveConvertOptions` sanitizes the metadata that never passes through
 * it (`options.ts`), so the whole `convert()` path is already safe. But
 * `renderDocument` is exported and documented as one of "both halves of the
 * pipeline… for callers who want to inspect the document in between", and a
 * model a *consumer* built — from a database row, another AST, a hand-written
 * literal — has been through neither.
 *
 * A single `U+0001` in such a model produces a `word/document.xml` that no
 * conforming parser will read, which in practice means Word refuses to open the
 * file at all and the caller gets no error of any kind. So every string the
 * renderer is about to hand to `docx` goes through here first. See
 * `src/xml-text.ts` for the character classes involved and why there is no
 * escape for them.
 *
 * The check is cheap by construction: {@link sanitizeXmlText} returns the
 * *identical* string when there is nothing to replace, so the common case is
 * one scan, no allocation, and a reference comparison.
 */

import { sanitizeXmlText } from "../xml-text.js";
import type { RenderContext } from "./context.js";

/**
 * Makes `value` safe to write into an OOXML part, reporting it if it was not.
 *
 * @param value - Any string bound for a `<w:t>`, an attribute or a core property.
 * @param ctx - Render context, for the warning channel.
 * @returns `value` unchanged, or a same-length string with every
 *   unrepresentable character replaced by `U+FFFD`.
 */
export function safeText(value: string, ctx: RenderContext): string {
  const clean = sanitizeXmlText(value);
  if (clean !== value) {
    // Once per document: a pasted terminal capture can carry thousands of
    // these, and one line telling the caller their model is unsanitized is
    // worth more than a thousand telling them so repeatedly.
    ctx.warnOnce(
      "unrepresentable-character",
      "the model contained characters XML 1.0 cannot represent (a C0 control other than " +
        "tab/LF/CR, an unpaired surrogate, or U+FFFE/U+FFFF); each was replaced with U+FFFD, " +
        "because writing one verbatim produces a .docx Word refuses to open",
    );
  }
  return clean;
}
