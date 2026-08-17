# dev.to article

> **Draft. Do not post until every gate in [`README.md`](README.md) is closed.**
> Post on **Day 7**. The article links the repo and shows `npm install`, so
> gates 1 and 2 both have to be shut first.

**Notes for the human before publishing**

- The front matter below sets `published: false` on purpose. Paste, save as a
  dev.to draft, read it once in the preview, then flip it.
- dev.to allows **four tags maximum**. Four are set.
- `cover_image` is a `TODO(human)` — the hero GIF frame from the root README's
  `TODO(hero)` block would work, exported as a still.
- `canonical_url`: only set it if you publish this on `ksprtech.com/blog` first
  and want that to be the canonical. If dev.to is the only home, **delete the
  line** rather than pointing it at the tool page — a canonical to a different
  page is not a cross-link, it is a deindex request.
- Every code block below is real code from the repo, not illustrative
  pseudocode. If you edit the source, edit the article.

---

```yaml
---
title: "Native Word equations from TeX, in the browser: TeX → MathML → OMML"
published: false
description: "docx's Math builders cannot express a matrix. Here is the pipeline I built instead, the XML injection I found on the way, and the bug I still have not fixed."
tags: typescript, webdev, showdev, opensource
cover_image: TODO(human)
---
```

---

## The one-line version of the problem

I wanted `$x^2$` in a markdown document to arrive in Word as something you can
**click and edit in the equation editor**. Not a PNG of an equation.

That distinction sounds cosmetic and is not. A picture of an equation cannot be
corrected, searched, restyled, or read aloud by a screen reader. It is a dead
end in a document somebody is going to edit. Most converters that handle maths
at all hand you the picture, because the picture is enormously easier.

Word's native format for this is **OMML** — Office Math Markup Language,
`<m:oMath>` elements living inline in `word/document.xml`, no drawing and no
media part:

```xml
<m:oMath>
  <m:sSup>
    <m:e><m:r><m:t>x</m:t></m:r></m:e>
    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
  </m:sSup>
</m:oMath>
```

Getting from `x^2` to that, in a browser tab, turned out to be four hops and two
security problems. This is the write-up.

## Attempt zero: just use the library's Math builders

I build the `.docx` with [`docx`](https://www.npmjs.com/package/docx), which is
a genuinely good package and does have a maths API. So the obvious move is to
walk the parsed TeX and emit builder calls.

Here is every Math class `docx` 9.7.1 exports, straight out of its type
declarations:

```
Math                    MathAngledBrackets      MathCurlyBrackets
MathDegree              MathDenominator         MathFraction
MathFunction            MathFunctionName        MathFunctionProperties
MathIntegral            MathLimit               MathLimitLower
MathLimitUpper          MathNumerator           MathPreSubSuperScript
MathRadical             MathRadicalProperties   MathRoundBrackets
MathRun                 MathSquareBrackets      MathSubScript
MathSubSuperScript      MathSum                 MathSuperScript
```

Read that list for what is _not_ in it. There is no matrix. `grep -c Matrix` on
that file returns `0`. OMML has `<m:m>` for matrices and the builder API simply
does not reach it, so `\begin{pmatrix} a & b \\ c & d \end{pmatrix}` has no
expressible form. The n-ary situation is similar: `MathIntegral` and `MathSum`
exist, but they are two specific operators rather than the general `<m:nary>`,
so `\bigoplus` and `\oint` are out too. No accents, no bars, no equation arrays.

And even if the coverage were complete, I would still need to write a full
TeX-to-builder-tree compiler by hand, which is the actual work, and the builders
would only be the last five percent of it.

So: skip the builders, generate OMML directly, and find a way to push a raw
string into the document.

## Attempt one: which TeX renderer?

The instinct is KaTeX, because everybody uses KaTeX. It is the wrong tool here:
KaTeX renders to HTML plus a stylesheet. That is a _visual_ representation, and
there is no way back from a pile of positioned spans to a semantic maths tree.

[temml](https://temml.org) renders the same TeX to **MathML** instead. MathML is
a semantic tree, and — this is the part that makes the whole approach work — it
maps onto OMML element for element, because both formats are describing the same
structures. `<mfrac>` is `<m:f>`. `<msup>` is `<m:sSup>`. Somebody has already
done the mapping: [`mathml2omml`](https://www.npmjs.com/package/mathml2omml).

temml also renders without a DOM, which means the same code runs in a Web Worker
and on a server. KaTeX-to-HTML would have needed a DOM to even inspect.

So the pipeline is:

```
TeX ──temml──▶ MathML ──mathml2omml──▶ <m:oMath> string ──▶ word/document.xml
```

Four hops, two of them third-party, and every one of them a place where a
document can get corrupted.

Both engines are optional peer dependencies reached only by bare-specifier
`import()` from a single file, so a document with no equations never downloads a
TeX parser:

```ts
export const loadMathEngine: MathEngineLoader = async (): Promise<MathEngine> => {
  const [temml, mathml2omml] = await Promise.all([import("temml"), import("mathml2omml")]);
  const renderToString = temml.default.renderToString;
  const { mml2omml } = mathml2omml;
  // …
};
```

## Getting a raw XML string into `docx`

`docx` has an escape hatch for exactly this: `ImportedXmlComponent.fromXmlString`.
Using it directly produces a corrupt file, and the reason is a good one.

Under the hood it parses with `xml-js`, which returns a **document** node
wrapping your element. That document node has no name. The packer serialises it
anyway, so the file ends up containing a literal `<undefined>` element and Word
offers to repair the whole document. The fix is to unwrap one level:

```ts
/**
 * `ImportedXmlComponent.fromXmlString` cannot be used directly: it converts the
 * xml-js *document* node, whose name is `undefined`, so the packer emits a
 * literal `<undefined>` element and corrupts the file. Unwrapping `root[0]`
 * yields the real `m:oMath` component. `root` is `protected`, hence the
 * structural check plus cast — and `ImportedXmlComponent` is not a member of
 * the `ParagraphChild` union either, though the packer accepts it.
 */
export function importOmml(omml: string): ParagraphChild | null {
  try {
    const imported: unknown = ImportedXmlComponent.fromXmlString(omml);
    if (!hasXmlRoot(imported)) return null;
    const first = imported.root[0];
    if (typeof first !== "object" || first === null) return null;
    return first as unknown as ParagraphChild;
  } catch {
    return null;
  }
}
```

Two casts in nine lines, both load-bearing, both commented, and I am not
thrilled about either. `root` is `protected`, so the structural check is
standing in for a type the library does not export. This is the honest cost of
reaching past a library's public API, and if you do it, write the paragraph
explaining why next to the cast — future you will otherwise delete it.

## The part where a document text becomes executable markup

Here is the bug that made me stop and rewrite the whole stage.

`mathml2omml` builds its output by **string concatenation** and does not escape
anything on the way out. Worse, by default it _decodes_ the XML entities in its
MathML input and writes the decoded characters back raw.

Follow one document through that. A user types this in their markdown:

```
$<w:p><w:r><w:t>PWNED</w:t></w:r></w:p>$
```

temml does the right thing and escapes it into `&lt;w:p&gt;…`. Then
`mathml2omml` decodes it back to `<w:p>…` and concatenates it into the output
string. It lands in `word/document.xml` as **live WordprocessingML**. Ordinary
document text has become markup, in a file the user is about to email to
somebody.

The milder version of the same defect is just as bad in practice: any equation
containing `a \& b` or `a < b` produces ill-formed XML, which corrupts the whole
package rather than one equation.

The fix at the source is one option:

```ts
mathmlToOmml(mathml) {
  // disableDecode is load-bearing, not a tweak: without it the converter
  // decodes temml's escaped MathML and writes the result raw, which turns
  // `$<w:p>…</w:p>$` into live markup in word/document.xml.
  return mml2omml(mathml, { disableDecode: true });
}
```

temml's MathML is already correctly escaped, so leaving it alone is exactly
right.

But I did not want the integrity of my output part to depend on a
`disableDecode` flag in someone else's package continuing to mean what it means
today. So there is a second line: a validator that re-reads **every** generated
fragment before it is embedded, and refuses anything that is not a namespaced
OMML fragment with properly escaped text.

The allowlist is the interesting design decision:

- **Elements:** anything in the `m:` namespace, wholesale, plus exactly three
  WordprocessingML elements the converter legitimately emits for bold and italic
  runs — `w:rPr`, `w:b`, `w:i`. Nothing else. In particular no un-namespaced
  element and **no other `w:*` element**, because that is where everything
  dangerous in WordprocessingML lives: `w:drawing` with an external
  relationship, `w:fldSimple`, `w:hyperlink`, `w:altChunk`.
- **Why `m:*` wholesale** rather than an exact name list: the OMML namespace
  contains nothing but maths layout. No element in it can reference anything
  outside the part. So allowing the namespace is safe, and it means a converter
  upgrade that emits a new element keeps working instead of silently degrading
  every equation that uses it. Allowlist the property that matters, not the
  enumeration you happen to know about today.
- **Attributes:** `m:*`, `xml:space` with one of its two legal values, and the
  two namespace declarations — which must carry the exact expected URIs and may
  only appear on the root. Nothing in `w:` or `r:`, because `r:id` is how you
  point at an external target.
- **Text:** no raw `<`; every `&` must start a real character reference; every
  character must be one XML 1.0 can actually hold.

And the scan itself is a single left-to-right pass with no backtracking regex,
so its cost is linear in the length of the fragment and a hostile equation
cannot turn the validator into a denial of service. That is not theoretical
tidiness — a 40,000-character equation is a legal thing for someone to paste,
and it is in the abuse corpus.

## One bug I chose to repair rather than refuse

`mathml2omml` maps MathML's `mathvariant` onto OMML's `ST_Style` enumeration
(`p | b | i | bi`) and its lookup table has no entry for `normal`. temml emits
`mathvariant="normal"` for `\Gamma`, `\Omega`, `\mathrm{}` and
`\operatorname{}` — so a capital Greek letter produces the literal string:

```xml
<m:sty m:val="undefined"/>
```

Word treats an out-of-range enumeration value as a **corrupt part** and offers
to repair the entire document. One `\Gamma`, one repair prompt.

Refusing the equation would turn every capital Greek letter into literal TeX, so
I repair this one specific case instead:

```ts
const BROKEN_STYLE = /<m:sty m:val="undefined"\/>/g;

export function repairOmml(omml: string): string {
  return omml.includes('m:val="undefined"')
    ? omml.replace(BROKEN_STYLE, '<m:sty m:val="p"/>')
    : omml;
}
```

The reason this is a repair and not a guess: `normal` means upright, OMML spells
upright `p` for plain, and the mapping is exact. Any _other_ attribute that
arrives as `"undefined"` is refused by the validator, because there I would only
be guessing. That line — repair what you can prove, refuse what you would have
to guess — is the one I would keep from this whole project.

## Nothing in the maths path is allowed to throw

A `.docx` must not be lost because somebody typed `\frac{a`. There is no
shortage of ways for TeX to go wrong, and every one of them lands on the same
rung:

| input                                | what happens                                     |
| ------------------------------------ | ------------------------------------------------ |
| `\frac{a`                            | temml throws `ParseError` → `tex-invalid`        |
| `\undefinedmacro`                    | `ParseError` → `tex-invalid`                     |
| `\def\a{\a}\a`                       | `ParseError` (`maxExpand`) → `tex-invalid`       |
| 1000-deep nested `\frac`             | `RangeError`, call stack → `tex-invalid`         |
| a 100 kB equation                    | converts, in about 200 ms                        |
| a 300 kB equation                    | over `maxLength` → `tex-too-large`, never parsed |
| `<w:p><w:r><w:t>x</w:t></w:r></w:p>` | escaped, kept escaped, re-checked → plain text   |

The degradation ladder has three rungs and you can only ever move down it:

```
picture ──▶ native equation (OMML) ──▶ the TeX source as literal text
```

The bottom rung is what the renderer already does for an unconverted equation,
so the document always gets produced and the reader always sees _something_ —
worst case, the TeX they typed. Every failure also emits a warning with a code,
a severity and the offending source, so the UI can tell them.

One more small thing that paid for itself: `mathml2omml` writes
`console.warn("Type not supported: mpadded")` straight to the host's console for
every MathML element it has no OMML for. Since it is synchronous — and
JavaScript is not re-entrant — swapping `console.warn` for the duration of the
call is safe, and turns that noise into a proper diagnostic on the library's own
warning channel. Results are also memoised on `(tex, displayMode)`, so a symbol
repeated forty times in a table is parsed once and warned about once.

## The bug I have not fixed

Being straight about this one, because it is visible in every document that
contains an integral.

`$\int_0^1 x^2 dx$` produces an `<m:nary>` whose body `<m:e/>` is **empty**, with
the integrand following as a sibling. So Word and LibreOffice draw the dotted
"empty slot" placeholder box after the integral sign. The equation is complete
and fully editable; the box is cosmetic and ugly.

The cause is upstream and it is a seam problem rather than anyone's bug: temml
emits `<msubsup>` plus following siblings, which is a perfectly correct MathML
encoding, and `mathml2omml` walks it element by element with nothing to hoist
into `<m:e>`. Fixing it needs a semantic rewrite pass over the generated OMML —
find the n-ary, pull the following siblings into its body — which is a real
piece of work on a tree I do not otherwise touch, and I have not done it.

If you take one thing from this section: when you chain two converters that were
written independently, the bugs do not live in either of them. They live in the
assumption each one makes about the other's output shape, and no amount of
testing either in isolation will find them.

## Five things I would tell you before you try this

1. **Check the escape hatch's failure mode before you build on it.** `docx`'s
   raw-XML import works fine; it just needed one unwrap that nothing documents.
2. **A converter that concatenates strings will eventually emit your user's
   input as markup.** Assume it, and validate the output rather than trusting
   the input path.
3. **Allowlist the property, not the enumeration.** `m:*` is safe because
   nothing in that namespace can reference anything outside the part. That
   reasoning survives a dependency upgrade; a hard-coded element list does not.
4. **Repair what you can prove; refuse what you would have to guess.** Both are
   defensible. Guessing quietly is not.
5. **Never throw away the document.** Every failure in this path degrades to the
   TeX source as visible text plus a warning. Nobody has ever been glad a
   converter silently deleted something.

## The honest state of it

This ships in a browser-based markdown → `.docx` converter I built. The whole
conversion runs in a Web Worker in the tab; the converter makes no network
requests. (The hosted page is on a normal marketing site that loads Google
Analytics, so I will not tell you the _page_ makes none — but your document is
never in a request.)

And the caveat that belongs at the end of an article like this one, since the
whole subject has been getting XML exactly right: **not one of these equations
has ever been looked at inside Microsoft Word.** What is automated is a headless
LibreOffice pass over every artefact the test suite produces — guarded by a
positive control on the import filter and a negative control on a corrupt
fixture, so the check cannot pass vacuously. That establishes a well-formed
container that one real implementation reads end to end. It establishes nothing
whatsoever about whether an `<m:nary>` typesets the way I think it does on a
Windows machine. Those are different sentences and I have been careful not to
let the first quietly become the second; the repo carries a fidelity matrix in
which every cell still reads "awaiting manual verification".

If you have Word and ten minutes, that is the most useful contribution anyone
could make.

- Tool: <https://ksprtech.com/tools/markdown-to-word>
- Source, MIT: <https://github.com/kspr-technologies/downword>
- The files this article is about: `packages/core/src/math/` — `engine.ts`,
  `convert.ts`, `omml.ts`.

```sh
npm install downword
```

```ts
import { convertWithMath } from "downword/plugins/math";

const bytes = await convertWithMath("Euler: $e^{i\\pi} + 1 = 0$\n", {
  math: { math: "omml", onWarning: (warning) => console.warn(warning.message) },
});
```
