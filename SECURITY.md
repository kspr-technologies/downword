# Security Policy

## Supported versions

downword is pre-1.0. Only the latest published version of each package receives
security fixes.

| Package                  | Supported           |
| ------------------------ | ------------------- |
| `@ksprtech/downword`     | latest release only |
| `@ksprtech/downword-cli` | latest release only |

## Reporting a vulnerability

**Please do not open a public issue for security problems.**

Report privately through GitHub:

1. Go to <https://github.com/kspr-technologies/downword/security/advisories/new>
2. Describe the issue, the affected version, and how to reproduce it.

If you cannot use GitHub Security Advisories, email **security@kspr.tech** with
`[downword]` in the subject line.

### What to expect

| Stage              | Target                                  |
| ------------------ | --------------------------------------- |
| Acknowledgement    | within 3 business days                  |
| Initial assessment | within 10 business days                 |
| Fix or mitigation  | depends on severity                     |
| Public disclosure  | after a fix ships, coordinated with you |

We will credit you in the advisory unless you ask us not to.

## Threat model

downword converts untrusted markdown into an OOXML document. The interesting
risks are therefore:

- **Malicious markdown → malicious document.** Anything that could make the
  generated `.docx` execute code, fetch remote content on open, or embed an
  external reference the user did not ask for (remote images, `INCLUDETEXT`
  fields, OLE objects, macros).
- **Injection into OOXML.** Markdown content that escapes its XML context and
  injects arbitrary parts, relationships, or field codes.
- **Denial of service.** Input that makes conversion hang or exhaust memory
  (pathological nesting, catastrophic backtracking in a regex, zip bombs in
  embedded assets).
- **Privacy.** downword's core promise is that nothing leaves the browser. Any
  code path that performs a network request from the library is a security bug,
  not a feature request.
- **Supply chain.** Releases are published from CI with npm provenance
  (`id-token: write`, `npm publish --provenance`). A published artifact that
  does not match the tagged commit is a security issue.

Out of scope: vulnerabilities in Microsoft Word or LibreOffice themselves, and
issues that require the attacker to already control the machine running
downword.
