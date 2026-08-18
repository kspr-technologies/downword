/**
 * The gate between a third-party XML *generator* and our OOXML package.
 *
 * ## Why this file exists
 *
 * `mathml2omml` builds its output by string concatenation and does not escape
 * anything on the way out. Two consequences, both reproducible against 0.5.0:
 *
 * 1. By default it **decodes** the XML entities in its MathML input and writes
 *    the decoded characters back raw. `$<w:p><w:r><w:t>PWNED</w:t></w:r></w:p>$`
 *    is escaped by `temml` into `&lt;w:p&gt;…`, decoded by `mathml2omml` back
 *    into `<w:p>…`, and lands in `word/document.xml` as **live WordprocessingML**.
 *    That is an XML injection from ordinary document text.
 * 2. Text that merely *contains* `&` or `<` (`a \& b`, `a < b`) produces
 *    ill-formed XML, which corrupts the package rather than one equation.
 *
 * `{ disableDecode: true }` (see `convert.ts`) fixes both at the source, because
 * `temml`'s MathML is already correctly escaped and leaving it alone is exactly
 * right. This file is the second line: it re-reads the generated string and
 * refuses anything that is not a namespaced OMML fragment with escaped text, so
 * a regression in the converter, or a future version that escapes differently,
 * cannot reach the package. `checkOmml` runs on **every** equation, not just
 * suspicious ones.
 *
 * ## What is allowed
 *
 * - Elements: anything in the OMML namespace (`m:*`), plus the three
 *   WordprocessingML elements the converter emits for bold/italic runs
 *   (`w:rPr`, `w:b`, `w:i`). Nothing else — in particular no un-namespaced
 *   element (`<script>`, `<b>`) and no other `w:*` element, which is where
 *   everything dangerous in WordprocessingML lives (`w:drawing` with an
 *   external relationship, `w:fldSimple`, `w:hyperlink`, `w:altChunk`).
 *   The `m:*` namespace is allowed wholesale rather than by an exact name list
 *   because it holds nothing but math layout — no reference to anything outside
 *   the part — so a converter upgrade that adds an element stays valid instead
 *   of silently degrading every equation that uses it.
 * - Attributes: anything in the `m:` namespace (`m:val`, `m:pos`), `xml:space`
 *   with one of its two legal values, and the two namespace declarations, which
 *   must carry the exact OMML/WML namespace URIs and may only appear on the
 *   root. Nothing else — in particular nothing in the `w:`/`r:` namespaces,
 *   which is where a relationship id pointing at an external target would go.
 * - Text: no `<`; every `&` starts a predefined or numeric character reference;
 *   every character is one XML 1.0 can carry.
 *
 * The scan is a single left-to-right pass with no backtracking regex, so its
 * cost is linear in the length of the OMML and a hostile equation cannot turn
 * it into a denial of service.
 */

import { sanitizeXmlText } from "../xml-text.js";

/** The `m:` namespace `mathml2omml` declares on its root. */
export const OMML_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** The `w:` namespace `mathml2omml` declares on its root. */
export const WML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The only WordprocessingML elements an OMML fragment may contain. */
const ALLOWED_W_ELEMENTS: ReadonlySet<string> = new Set(["w:rPr", "w:b", "w:i"]);

/** `xml:space` is an XML-defined attribute with exactly two legal values. */
const ALLOWED_XML_SPACE: ReadonlySet<string> = new Set(["preserve", "default"]);

/** An `m:`-namespaced element or attribute name, and nothing cleverer. */
const M_NAME = /^m:[A-Za-z][A-Za-z0-9]*$/;

/** One character reference, anchored: the five predefined names plus numeric forms. */
const ENTITY = /&(?:amp|lt|gt|quot|apos|#[0-9]{1,7}|#x[0-9A-Fa-f]{1,6});/y;

/**
 * The one `mathml2omml` defect worth repairing rather than refusing.
 *
 * Its `STYLES` table maps MathML's `mathvariant` onto OMML's `ST_Style`
 * enumeration (`p | b | i | bi`) but has no entry for `normal` — so
 * `mathvariant="normal"`, which is what `temml` emits for `\Gamma`, `\Omega`,
 * `\mathrm{}` and `\operatorname{}`, produces the literal
 * `<m:sty m:val="undefined"/>`. Word treats an out-of-range enumeration value
 * as a corrupt part and offers to repair the whole document, so leaving it
 * alone is not an option, and refusing the equation would turn every capital
 * Greek letter into literal TeX.
 *
 * `normal` means upright, and OMML spells upright `p` (plain), so the repair is
 * exact rather than a guess. Anything else that arrives as `"undefined"` is
 * refused by {@link checkOmml} instead, because we would only be guessing.
 */
const BROKEN_STYLE = /<m:sty m:val="undefined"\/>/g;

/** An empty n-ary operand, which is what {@link hoistNaryOperands} repairs. */
const EMPTY_NARY_OPERAND = "<m:e/></m:nary>";

/**
 * Returns the index just past the element starting at `open`, or -1.
 *
 * Depth counting over `m:`-namespaced tags only, which is sound here because
 * {@link checkOmml} has already refused anything that is not `m:` (plus the
 * three allowed `w:` run-property elements, which are always balanced inside
 * an `m:r`). Self-closing tags never change depth.
 */
function elementEnd(xml: string, open: number): number {
  if (xml[open] !== "<") return -1;
  let depth = 0;
  let i = open;
  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    if (lt === -1) return -1;
    const gt = xml.indexOf(">", lt);
    if (gt === -1) return -1;
    const selfClosing = xml[gt - 1] === "/";
    const closing = xml[lt + 1] === "/";
    if (!selfClosing && !closing) depth += 1;
    else if (closing) depth -= 1;
    if (depth === 0) return gt + 1;
    i = gt + 1;
  }
  return -1;
}

/**
 * Moves an n-ary operator's operand inside the operator.
 *
 * MathML models `\int_0^1 x^2 dx` as siblings in one `<mrow>` — the `<mo>` and
 * then the integrand — because the operator does not contain its operand.
 * OMML models the same expression the other way round: `<m:nary>` *contains*
 * its operand in `<m:e>`. `mathml2omml` maps the operator faithfully and then
 * has nothing to put in `<m:e>`, so it emits `<m:e/>` and leaves the integrand
 * as the n-ary's next sibling.
 *
 * That is not merely cosmetic. `m:e` is required, and an empty one is a hole in
 * the equation: LibreOffice draws a placeholder box, Word shows an empty slot,
 * and the integral semantically has no integrand. Verified in a rendered PDF —
 * `∫₀¹ □ x²dx` — which is why this is repaired rather than tolerated.
 *
 * The repair hoists exactly ONE following element. For the shapes temml
 * produces that is the whole operand and nothing more: `∫₀¹(x²)dx = 1/3` and
 * `∑ₖ₌₁ⁿ(k²) = …` are both what a human would have typed in Word's editor.
 * Taking more would be a guess about where an integrand ends, and there is no
 * general answer to that — `dx` is a delimiter by convention, not by markup.
 * An n-ary with nothing after it is left alone: there is nothing to hoist.
 */
function hoistNaryOperands(omml: string): string {
  let out = omml;
  let from = 0;
  for (;;) {
    const hole = out.indexOf(EMPTY_NARY_OPERAND, from);
    if (hole === -1) return out;
    const afterNary = hole + EMPTY_NARY_OPERAND.length;
    const end = elementEnd(out, afterNary);
    if (end === -1) {
      // Nothing balanced follows: leave the hole rather than corrupt the tree.
      from = afterNary;
      continue;
    }
    const operand = out.slice(afterNary, end);
    out = out.slice(0, hole) + "<m:e>" + operand + "</m:e></m:nary>" + out.slice(end);
    from = hole + "<m:e>".length + operand.length + "</m:e></m:nary>".length;
  }
}

/**
 * Applies {@link BROKEN_STYLE} and {@link hoistNaryOperands}.
 *
 * Returns the input unchanged when there is nothing to do.
 */
export function repairOmml(omml: string): string {
  const styled = omml.includes('m:val="undefined"')
    ? omml.replace(BROKEN_STYLE, '<m:sty m:val="p"/>')
    : omml;
  return styled.includes(EMPTY_NARY_OPERAND) ? hoistNaryOperands(styled) : styled;
}

/** What {@link checkOmml} decided. */
export type OmmlCheck =
  | {
      readonly ok: true;
      /** `<m:oMath/>` with no children — a valid fragment that renders as nothing. */
      readonly empty: boolean;
    }
  | {
      readonly ok: false;
      /** Why it was refused, phrased for a warning message. */
      readonly reason: string;
    };

function refuse(reason: string): OmmlCheck {
  return { ok: false, reason };
}

/** Whether a UTF-16 unit ends a tag name or an attribute name. */
function isNameEnd(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

/**
 * Validates a text or attribute-value run: no `<`, only real entities, only
 * characters XML 1.0 can hold.
 */
function checkChars(value: string, where: string): string | null {
  if (value.includes("<")) return `${where} contains a raw "<"`;
  if (sanitizeXmlText(value) !== value) {
    return `${where} contains a character XML cannot represent`;
  }

  let at = value.indexOf("&");
  while (at !== -1) {
    ENTITY.lastIndex = at;
    if (!ENTITY.test(value)) return `${where} contains an unescaped "&"`;
    at = value.indexOf("&", ENTITY.lastIndex);
  }
  return null;
}

/**
 * Decides whether a generated OMML string may be embedded.
 *
 * @param omml - Whatever the MathML -> OMML converter returned.
 * @returns `{ ok: true }` and whether the equation is empty, or the reason it
 *   was refused. Never throws.
 */
export function checkOmml(omml: string): OmmlCheck {
  if (typeof omml !== "string" || omml === "") return refuse("the converter returned nothing");
  // mathml2omml stringifies a document node with no name as the literal text
  // "undefined" when its input contains no <math> element.
  if (!omml.startsWith("<m:oMath")) return refuse("the converter did not return an <m:oMath>");

  const stack: string[] = [];
  let closedRoot = false;
  let elements = 0;
  let index = 0;

  while (index < omml.length) {
    const lt = omml.indexOf("<", index);
    const text = omml.slice(index, lt === -1 ? omml.length : lt);

    if (text !== "") {
      if (stack.length === 0 && text.trim() !== "") {
        return refuse("text outside the <m:oMath> root");
      }
      const bad = checkChars(text, "the equation text");
      if (bad !== null) return refuse(bad);
    }
    if (lt === -1) break;

    if (omml.startsWith("<!", lt) || omml.startsWith("<?", lt)) {
      return refuse("a comment, doctype or processing instruction");
    }

    const closing = omml.startsWith("</", lt);
    let cursor = lt + (closing ? 2 : 1);
    const nameStart = cursor;
    while (cursor < omml.length) {
      const code = omml.charCodeAt(cursor);
      if (isNameEnd(code) || code === 0x2f /* / */ || code === 0x3e /* > */) break;
      cursor += 1;
    }
    const name = omml.slice(nameStart, cursor);
    if (name === "") return refuse("an element with no name");
    if (!M_NAME.test(name) && !ALLOWED_W_ELEMENTS.has(name)) {
      return refuse(`a <${name}> element, which is not OMML`);
    }

    if (closing) {
      while (cursor < omml.length && isNameEnd(omml.charCodeAt(cursor))) cursor += 1;
      if (omml.charCodeAt(cursor) !== 0x3e) return refuse(`a malformed </${name}>`);
      if (stack.pop() !== name) return refuse(`</${name}> does not close the open element`);
      if (stack.length === 0) closedRoot = true;
      index = cursor + 1;
      continue;
    }

    if (closedRoot) return refuse("a second element after the <m:oMath> root closed");
    const isRoot = stack.length === 0;
    if (isRoot && name !== "m:oMath") return refuse(`a <${name}> root instead of <m:oMath>`);
    elements += 1;

    // Attributes.
    let selfClosing = false;
    for (;;) {
      while (cursor < omml.length && isNameEnd(omml.charCodeAt(cursor))) cursor += 1;
      const code = omml.charCodeAt(cursor);
      if (code === 0x3e /* > */) {
        cursor += 1;
        break;
      }
      if (code === 0x2f /* / */) {
        if (omml.charCodeAt(cursor + 1) !== 0x3e) return refuse(`a malformed <${name}>`);
        selfClosing = true;
        cursor += 2;
        break;
      }
      if (Number.isNaN(code)) return refuse(`an unterminated <${name}>`);

      const attrStart = cursor;
      while (cursor < omml.length) {
        const at = omml.charCodeAt(cursor);
        if (isNameEnd(at) || at === 0x3d /* = */ || at === 0x2f || at === 0x3e) break;
        cursor += 1;
      }
      const attribute = omml.slice(attrStart, cursor);
      if (omml.charCodeAt(cursor) !== 0x3d) return refuse(`a bare attribute on <${name}>`);
      cursor += 1;
      if (omml.charCodeAt(cursor) !== 0x22 /* " */) {
        return refuse(`an unquoted value for ${attribute} on <${name}>`);
      }
      cursor += 1;
      const valueStart = cursor;
      const valueEnd = omml.indexOf('"', cursor);
      if (valueEnd === -1) return refuse(`an unterminated value for ${attribute}`);
      const value = omml.slice(valueStart, valueEnd);
      cursor = valueEnd + 1;

      if (attribute === "xmlns:m" || attribute === "xmlns:w") {
        if (!isRoot) return refuse(`a namespace declaration on <${name}>`);
        const expected = attribute === "xmlns:m" ? OMML_NAMESPACE : WML_NAMESPACE;
        if (value !== expected) return refuse(`${attribute}="${value}", which is not OMML's`);
      } else if (attribute === "xml:space") {
        if (!ALLOWED_XML_SPACE.has(value)) return refuse(`xml:space="${value}"`);
      } else if (M_NAME.test(attribute)) {
        // A value the converter failed to map. Every m:val is an enumeration or
        // a single character, so the string "undefined" is never legitimate and
        // an out-of-range enumeration is what makes Word declare a part corrupt.
        if (value === "undefined") return refuse(`${attribute}="undefined"`);
        const bad = checkChars(value, `the ${attribute} value`);
        if (bad !== null) return refuse(bad);
      } else {
        return refuse(`a ${attribute} attribute on <${name}>`);
      }
    }

    if (!selfClosing) stack.push(name);
    else if (isRoot) closedRoot = true;
    index = cursor;
  }

  if (stack.length > 0) return refuse(`an unclosed <${stack[stack.length - 1] ?? "?"}>`);
  if (!closedRoot) return refuse("no closed <m:oMath> root");
  return { ok: true, empty: elements <= 1 };
}
