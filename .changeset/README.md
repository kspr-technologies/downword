# Changesets

This folder is managed by [changesets](https://github.com/changesets/changesets).

Every user-visible change needs a changeset. From the repository root:

```sh
pnpm changeset
```

Pick the affected packages, pick a bump type (`patch` / `minor` / `major`), and
write the entry as release notes a user would read — not as a commit message.

Releases are automated. `.github/workflows/release.yml` turns the changesets on
`main` into a "Version Packages" PR; merging that PR publishes to npm with
provenance. See [CONTRIBUTING.md](../CONTRIBUTING.md#7-commit-changesets-and-releases).
