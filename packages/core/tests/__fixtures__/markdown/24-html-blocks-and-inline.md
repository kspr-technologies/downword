Some inline HTML: a line<br>break, H<sub>2</sub>O, x<sup>2</sup>, and
<mark>highlighted</mark> text with <span style="color:red">a span</span>.

<div align="center">
  <img src="logo.png" width="200" alt="Logo">
  <p>A centred block, which LLMs emit whenever you ask for centring.</p>
</div>

<details>
<summary>Click to expand</summary>

Markdown **inside** an HTML block is still parsed when there is a blank line.

</details>

A raw table, because the model forgot GFM tables exist:

<table>
  <tr><th>a</th><th>b</th></tr>
  <tr><td>1</td><td>2</td></tr>
</table>

An HTML comment: <!-- this should not be visible -->

A self-closing tag: <hr />

An unclosed tag at the end: <em>dangling
