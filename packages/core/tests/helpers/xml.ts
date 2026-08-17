/**
 * Structural probes for generated OOXML.
 *
 * A snapshot alone is a weak test: it locks in whatever the renderer produced,
 * including a heading rendered with direct formatting and no `<w:pStyle>` at
 * all. These helpers let a test assert *positively* — "this paragraph carries
 * `Heading2`", "this list paragraph joins numbering instance 3, level 1" —
 * so the snapshot only has to guard against unintended drift.
 *
 * The parsing is regex-based, which would be indefensible for arbitrary XML but
 * is fine here: the input is always output we generated ourselves in this same
 * process, with no CDATA, no comments and no namespace juggling.
 */

/** One `<w:p …>…</w:p>` element, verbatim. */
export type ParagraphXml = string;

/** Every paragraph in a `word/document.xml`, in document order. */
export function paragraphs(documentXml: string): ParagraphXml[] {
  return [...documentXml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:p\s*\/>/g)].map((m) => m[0]);
}

/** Every `<w:tbl>` element in a `word/document.xml`. */
export function tables(documentXml: string): string[] {
  return [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m) => m[0]);
}

/** Every `<w:tr …>` element of a table. */
export function rows(tableXml: string): string[] {
  return [...tableXml.matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
}

/** Every `<w:tc>` element of a row. */
export function cells(rowXml: string): string[] {
  return [...rowXml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((m) => m[0]);
}

/** The `w:val` of an element's `<w:pStyle>`, or `null` if it has none. */
export function pStyle(xml: ParagraphXml): string | null {
  return /<w:pStyle w:val="([^"]*)"\/>/.exec(xml)?.[1] ?? null;
}

/** Every `<w:rStyle w:val>` inside an element, in order. */
export function rStyles(xml: string): string[] {
  return [...xml.matchAll(/<w:rStyle w:val="([^"]*)"\/>/g)].map((m) => m[1] ?? "");
}

/** The numbering reference of a paragraph: its level and concrete `numId`. */
export function numPr(xml: ParagraphXml): { ilvl: number; numId: number } | null {
  const match = /<w:numPr><w:ilvl w:val="(\d+)"\/><w:numId w:val="(\d+)"\/><\/w:numPr>/.exec(xml);
  if (match === null) return null;
  return { ilvl: Number(match[1]), numId: Number(match[2]) };
}

/** Concatenated visible text of an element (`<w:t>` contents only). */
export function textOf(xml: string): string {
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((m) => m[1] ?? "")
    .join("")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"');
}

/** Finds the first paragraph whose text equals `text`. */
export function paragraphWithText(documentXml: string, text: string): ParagraphXml {
  const found = paragraphs(documentXml).find((p) => textOf(p) === text);
  if (found === undefined) {
    throw new Error(
      `no paragraph with text ${JSON.stringify(text)}; saw ${JSON.stringify(
        paragraphs(documentXml).map(textOf),
      )}`,
    );
  }
  return found;
}

/** The `<w:ind …/>` attributes of a paragraph, as raw strings. */
export function indent(xml: ParagraphXml): Record<string, string> | null {
  const match = /<w:ind ([^/>]*)\/>/.exec(xml);
  if (match === null) return null;
  const out: Record<string, string> = {};
  for (const attr of (match[1] ?? "").matchAll(/([\w:]+)="([^"]*)"/g)) {
    out[attr[1] ?? ""] = attr[2] ?? "";
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* styles.xml                                                                  */
/* -------------------------------------------------------------------------- */

/** A single `<w:style>` element from `word/styles.xml`, by style id. */
export function styleById(stylesXml: string, id: string): string | null {
  const pattern = new RegExp(`<w:style [^>]*w:styleId="${id}">[\\s\\S]*?</w:style>`);
  return pattern.exec(stylesXml)?.[0] ?? null;
}

/** Every style id declared in `word/styles.xml`. */
export function styleIds(stylesXml: string): string[] {
  return [...stylesXml.matchAll(/w:styleId="([^"]*)"/g)].map((m) => m[1] ?? "");
}

/* -------------------------------------------------------------------------- */
/* numbering.xml                                                               */
/* -------------------------------------------------------------------------- */

/** A concrete numbering instance: which abstract it uses and where it starts. */
export interface ConcreteNum {
  readonly numId: number;
  readonly abstractNumId: number;
  readonly startOverride: number | null;
}

/** Every `<w:num>` in `word/numbering.xml`. */
export function concreteNums(numberingXml: string): ConcreteNum[] {
  return [...numberingXml.matchAll(/<w:num w:numId="(\d+)">([\s\S]*?)<\/w:num>/g)].map((match) => {
    const body = match[2] ?? "";
    const abstract = /<w:abstractNumId w:val="(\d+)"\/>/.exec(body)?.[1] ?? "-1";
    const start = /<w:startOverride w:val="(-?\d+)"\/>/.exec(body)?.[1];
    return {
      numId: Number(match[1]),
      abstractNumId: Number(abstract),
      startOverride: start === undefined ? null : Number(start),
    };
  });
}

/** A single `<w:abstractNum>` element, by abstract id. */
export function abstractNum(numberingXml: string, abstractNumId: number): string | null {
  const pattern = new RegExp(
    `<w:abstractNum w:abstractNumId="${abstractNumId}"[^>]*>[\\s\\S]*?</w:abstractNum>`,
  );
  return pattern.exec(numberingXml)?.[0] ?? null;
}

/** One level of an abstract numbering definition. */
export interface NumberingLevel {
  readonly ilvl: number;
  readonly start: number | null;
  readonly format: string | null;
  readonly text: string | null;
  readonly indentLeft: number | null;
  readonly hanging: number | null;
  readonly font: string | null;
}

/** Parses the `<w:lvl>` children of an `<w:abstractNum>`. */
export function numberingLevels(abstractXml: string): NumberingLevel[] {
  return [...abstractXml.matchAll(/<w:lvl w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g)].map(
    (match) => {
      const body = match[2] ?? "";
      const number = (pattern: RegExp): number | null => {
        const found = pattern.exec(body)?.[1];
        return found === undefined ? null : Number(found);
      };
      return {
        ilvl: Number(match[1]),
        start: number(/<w:start w:val="(-?\d+)"\/>/),
        format: /<w:numFmt w:val="([^"]*)"\/>/.exec(body)?.[1] ?? null,
        text: /<w:lvlText w:val="([^"]*)"\/>/.exec(body)?.[1] ?? null,
        indentLeft: number(/<w:ind[^/>]*w:left="(\d+)"/),
        hanging: number(/<w:ind[^/>]*w:hanging="(\d+)"/),
        font: /<w:rFonts[^/>]*w:ascii="([^"]*)"/.exec(body)?.[1] ?? null,
      };
    },
  );
}

/** Looks a concrete `numId` up and returns the abstract definition it points at. */
export function abstractForNumId(numberingXml: string, numId: number): string {
  const concrete = concreteNums(numberingXml).find((num) => num.numId === numId);
  if (concrete === undefined) throw new Error(`no <w:num w:numId="${numId}"> in numbering.xml`);
  const abstract = abstractNum(numberingXml, concrete.abstractNumId);
  if (abstract === null) {
    throw new Error(`numId ${numId} points at missing abstract ${concrete.abstractNumId}`);
  }
  return abstract;
}
