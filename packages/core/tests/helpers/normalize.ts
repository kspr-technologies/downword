/**
 * Normalisation helpers for golden (snapshot) tests.
 *
 * A `.docx` is a zip of XML parts. Several of those parts contain values that
 * change between runs, between machines, or between `docx` releases. Snapshots
 * must be byte-stable across runs, so everything nondeterministic is replaced
 * with a fixed token here **before** it reaches the snapshot.
 *
 * Anything you add to this list must be genuinely nondeterministic - do not use
 * it to paper over real output changes, which are exactly what the golden test
 * exists to catch.
 */

/** Replacement token used for every scrubbed value. */
const TOKEN = "__NORMALIZED__";

interface Scrubber {
  readonly what: string;
  readonly pattern: RegExp;
  readonly replacement: string;
}

const SCRUBBERS: readonly Scrubber[] = [
  {
    what: "revision save ids (w:rsid, w:rsidR, w:rsidRDefault, ...)",
    pattern: /\s+w:rsid[A-Za-z]*="[^"]*"/g,
    replacement: "",
  },
  {
    what: "Word 2010+ paragraph/text ids",
    pattern: /\s+w14:(?:paraId|textId)="[^"]*"/g,
    replacement: "",
  },
  {
    what: "DrawingML anchor ids (regenerated per run)",
    pattern: /\s+wp14:(?:anchorId|editId)="[^"]*"/g,
    replacement: "",
  },
  {
    what: "numeric element ids (bookmarks, comments, footnotes, drawings)",
    pattern: /\s(w:id|id)="\d+"/g,
    replacement: ` $1="${TOKEN}"`,
  },
  {
    what: "ISO-8601 timestamps",
    pattern: /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?/g,
    replacement: TOKEN,
  },
  {
    // docx mints external-hyperlink relationship ids as "rId" + nanoid(21), so
    // they differ on every render of identical input. Image and header
    // relationships use sequential numeric ids and are deliberately NOT
    // scrubbed here - those have to stay stable.
    what: "nanoid-suffixed relationship ids (external hyperlinks)",
    pattern: /\br:id="rId[A-Za-z0-9_-]{21}"/g,
    replacement: `r:id="rId${TOKEN}"`,
  },
  {
    what: "nanoid-style ids emitted by the docx writer",
    pattern: /"[A-Za-z0-9_-]{21}"/g,
    replacement: `"${TOKEN}"`,
  },
  {
    what: "generator/version comments",
    pattern: /<!--[\s\S]*?-->/g,
    replacement: "",
  },
  {
    what: "sha1-named media parts",
    pattern: /\b[0-9a-f]{40}\b/g,
    replacement: TOKEN,
  },
];

/**
 * Scrubs nondeterministic values out of an OOXML part and reflows it to one
 * tag per line so snapshot diffs are readable.
 *
 * Splitting on `><` is safe: inside XML, literal `<` and `>` in text content
 * are always escaped as `&lt;` / `&gt;`, so a `><` sequence can only ever be a
 * tag boundary.
 */
export function normalizeOoxml(xml: string): string {
  let out = xml.replace(/^\uFEFF/, "");

  for (const { pattern, replacement } of SCRUBBERS) {
    out = out.replace(pattern, replacement);
  }

  return `${out.replace(/></g, ">\n<").trim()}\n`;
}

/** Human-readable list of what {@link normalizeOoxml} removes (used in a test). */
export function describeScrubbers(): readonly string[] {
  return SCRUBBERS.map((scrubber) => scrubber.what);
}
