Use `` ` `` to open a code span, and `` `` ` `` `` when the span itself needs a
double backtick.

A span containing a pipe: `grep 'a|b'`. A span that is only a pipe: `` | ``.

A span containing backticks: `` a ` b ``, and one containing three: ``` a `` b ```.

An empty-ish span: `` `` (that is a single space).

A span with a newline in the source, `which markdown-it
folds to a space`, and a span with **no** emphasis inside: `**not bold**`.

Code spans inside other marks: **`bold code`**, *`italic code`*,
~~`struck code`~~, [`linked code`](https://example.com), and
**[`bold linked code`](https://example.com)**.

Underscores inside code are literal: `snake_case_name`.
