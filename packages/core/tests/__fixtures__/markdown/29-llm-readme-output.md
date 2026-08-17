<h1 align="center">flux-cache</h1>

<p align="center">
  <a href="https://github.com/example/flux-cache/actions"><img src="https://img.shields.io/github/actions/workflow/status/example/flux-cache/ci.yml?branch=main" alt="CI"></a>
  <a href="https://www.npmjs.com/package/flux-cache"><img src="https://img.shields.io/npm/v/flux-cache.svg" alt="npm"></a>
</p>

> A tiny, dependency-free LRU cache with TTL support. **~1.2 kB gzipped.**

## Installation

```sh
npm install flux-cache
# or
pnpm add flux-cache
# or
yarn add flux-cache
```

## Quick start

```ts
import { FluxCache } from "flux-cache";

const cache = new FluxCache<string, User>({ max: 500, ttl: 60_000 });

cache.set("u_1", { id: "u_1", name: "Ada" });
cache.get("u_1"); // -> { id: "u_1", name: "Ada" }
cache.get("u_2"); // -> undefined
```

## API

### `new FluxCache(options)`

| Option | Type | Default | Description |
| ------ | ---- | ------- | ----------- |
| `max` | `number` | `1000` | Maximum number of entries before eviction |
| `ttl` | `number \| null` | `null` | Time-to-live in ms; `null` disables expiry |
| `onEvict` | `(key, value) => void` | — | Called for every evicted entry |
| `clock` | `() => number` | `Date.now` | Injectable clock, for tests |

### Methods

- `get(key)` — returns the value, or `undefined`. **Refreshes recency.**
- `peek(key)` — same, but *does not* refresh recency.
- `set(key, value, ttl?)` — per-entry `ttl` overrides the constructor's.
- `delete(key)` → `boolean`
- `clear()`
- `size` (getter) → `number`

## Benchmarks

Run on an M2 Pro, Node 22, 1e6 operations:

| Library | ops/sec | Relative |
|---------|--------:|---------:|
| `flux-cache` | 12,400,000 | **1.00x** |
| `lru-cache` | 9,800,000 | 0.79x |
| `quick-lru` | 11,100,000 | 0.90x |

<details>
<summary>Full benchmark output</summary>

```
flux-cache#get  x 12,412,003 ops/sec ±0.42% (94 runs sampled)
lru-cache#get   x  9,801,224 ops/sec ±0.61% (92 runs sampled)
quick-lru#get   x 11,102,881 ops/sec ±0.38% (95 runs sampled)
```

</details>

## Caveats

1. Keys are compared with `===`, so object keys must be identity-stable.
2. TTL uses a *lazy* sweep — an expired entry still occupies memory until it is
   read or evicted. If you need eager eviction, pass `sweepInterval`.
3. Not thread-safe across workers (there are no threads in JS, but there *are*
   `SharedArrayBuffer` users; this is not for them).

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md). TL;DR:

- [x] Fork and branch from `main`
- [x] `pnpm install && pnpm test`
- [ ] Add a changeset: `pnpm changeset`
- [ ] Open the PR against `main`

## License

MIT © [Example Corp](https://example.com)
