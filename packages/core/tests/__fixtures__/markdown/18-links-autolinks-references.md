An inline link: [the docs](https://example.com/docs).

An inline link with a title: [the docs](https://example.com/docs "Read this first").

An autolink: <https://example.com/auto> and an e-mail autolink
<hello@example.com>.

A bare URL that linkify picks up: https://bare.example.com/path?x=1#frag

A bare e-mail: support@example.com

A reference link: [see the guide][guide] and a collapsed one: [guide][].

A shortcut reference: [guide].

An intra-document anchor: [jump to setup](#setup).

A relative link: [./CONTRIBUTING.md](./CONTRIBUTING.md).

A link whose text is formatted: [**bold** and *italic* and `code`](https://example.com).

A link with a URL that needs escaping: [spaces](https://example.com/a%20b).

A dangerous link is dropped by markdown-it's link validator:
[click me](javascript:alert(1)).

[guide]: https://example.com/guide "The Guide"
