/**
 * The Node resolver: local files, plus the same capped HTTPS path.
 *
 * **This file is deliberately not exported from `./index.js`.** It is the only
 * module in `src/images/**` that imports `node:` builtins, and keeping it off
 * the barrel means a browser bundler never has to resolve them - the browser
 * entry's import graph is checked for exactly that in the test suite. Node
 * hosts (the CLI) import it by path:
 *
 * ```ts
 * import { createNodeImageResolver } from "@ksprtech/downword/…/images/node.js";
 * ```
 *
 * It adds three things to {@link createImageResolver}:
 *
 * 1. **Filesystem sources** - relative paths, absolute paths, and `file:` URLs,
 *    which is how `![](./diagram.png)` next to a markdown file is meant to work.
 * 2. **A base directory** the filesystem cannot escape. Converting an untrusted
 *    markdown file is otherwise an arbitrary-file-read: `![](/etc/passwd)` or
 *    `![](../../../.ssh/id_rsa)` would be embedded into the `.docx` and handed
 *    back to whoever supplied the document. Paths are resolved *and*
 *    `realpath`ed before the containment check, so a symlink pointing out of
 *    the tree does not slip through.
 * 3. **An SSRF guard.** A browser has the same-origin policy; a Node process
 *    has a route to `169.254.169.254` and to every service on `localhost`. When
 *    remote images are enabled, hosts that resolve into loopback, link-local,
 *    RFC1918 or CGNAT space are refused.
 *
 * Remote fetching still goes through `transport.ts`'s WHATWG `fetch` rather
 * than `node:https` - deliberately, so the 10 MiB cap, the 10 s abort, redirect
 * handling and the "never throws" contract are *the same code* in both
 * runtimes. A second, Node-only implementation of those caps would be a second
 * place for them to be subtly wrong. `fetch` has been stable in Node since v18,
 * and this package requires v20.
 */

import { lookup } from "node:dns/promises";
import { readFile, realpath, stat } from "node:fs/promises";
import { isAbsolute, resolve as resolvePath, sep } from "node:path";
import { fileURLToPath } from "node:url";

import type { ImageResolver } from "../render/types.js";
import {
  createImageResolver,
  type ImageLoadRequest,
  type ImageResolverOptions,
  type RemoteGuard,
} from "./resolver.js";
import { describeError } from "./transport.js";
import { bytesFailure, type ImageBytesResult } from "./types.js";

/** Options for {@link createNodeImageResolver}. */
export interface NodeImageResolverOptions extends Omit<
  ImageResolverOptions,
  "loaders" | "guardRemote"
> {
  /**
   * The directory relative paths resolve against, and cannot escape. Defaults
   * to `process.cwd()`. A CLI should set it to the markdown file's directory.
   */
  readonly baseDir?: string | undefined;
  /**
   * Let filesystem sources point outside {@link NodeImageResolverOptions.baseDir}.
   * Defaults to `false`. Only turn it on for documents you wrote yourself.
   */
  readonly allowOutsideBaseDir?: boolean | undefined;
  /** Read local files at all. Defaults to `true`. */
  readonly allowFilesystem?: boolean | undefined;
  /**
   * Refuse remote hosts that resolve into private address space. Defaults to
   * `true`.
   *
   * The check runs on the literal host and on every address `dns.lookup()`
   * returns for it. It is not airtight - the connection re-resolves the name,
   * so a DNS entry that flips between a public and a private address between
   * the two lookups (a "DNS rebinding" attack) can still win the race. Closing
   * that would mean resolving the address here and connecting to it directly
   * with a `Host` header, which means dropping `fetch` and hand-rolling
   * redirects and TLS SNI. The guard is a strong speed bump on the realistic
   * threat (a document that simply links to `http://localhost:9200/…`), not a
   * sandbox.
   */
  readonly blockPrivateNetwork?: boolean | undefined;
}

/** Whether a source looks like a Windows path (`C:\…`), which `new URL()` mis-parses as a scheme. */
function isWindowsPath(src: string): boolean {
  return /^[A-Za-z]:[\\/]/.test(src);
}

/** Strips the brackets IPv6 hostnames carry inside a URL. */
function bareHostname(hostname: string): string {
  return hostname.startsWith("[") && hostname.endsWith("]") ? hostname.slice(1, -1) : hostname;
}

/** Parses a dotted-quad IPv4 address into its octets. */
function parseIpv4(address: string): readonly number[] | null {
  const parts = address.split(".");
  if (parts.length !== 4) return null;
  const octets: number[] = [];
  for (const part of parts) {
    if (!/^\d{1,3}$/.test(part)) return null;
    const value = Number.parseInt(part, 10);
    if (value > 255) return null;
    octets.push(value);
  }
  return octets;
}

/** Whether an IPv4 address is anything other than ordinary public unicast space. */
function isPrivateIpv4(octets: readonly number[]): boolean {
  const [a = 0, b = 0] = octets;
  if (a === 0) return true; // "this network"
  if (a === 10) return true; // RFC1918
  if (a === 127) return true; // loopback
  if (a === 169 && b === 254) return true; // link-local, incl. cloud metadata
  if (a === 172 && b >= 16 && b <= 31) return true; // RFC1918
  if (a === 192 && b === 168) return true; // RFC1918
  if (a === 192 && b === 0) return true; // IETF protocol assignments
  if (a === 100 && b >= 64 && b <= 127) return true; // CGNAT
  if (a === 198 && (b === 18 || b === 19)) return true; // benchmarking
  if (a >= 224) return true; // multicast, reserved, broadcast
  return false;
}

/** Whether an address literal is loopback, link-local, unique-local or otherwise not public. */
export function isPrivateAddress(address: string): boolean {
  const host = bareHostname(address.trim().toLowerCase());
  if (host === "") return true;

  const ipv4 = parseIpv4(host);
  if (ipv4 !== null) return isPrivateIpv4(ipv4);

  if (!host.includes(":")) {
    // A name, not a literal. Only the well-known local ones are decidable here;
    // the rest are settled by the DNS lookup in the guard.
    return host === "localhost" || host.endsWith(".localhost") || host.endsWith(".local");
  }

  // IPv4-mapped and IPv4-compatible forms: ::ffff:127.0.0.1
  const lastColon = host.lastIndexOf(":");
  const tail = host.slice(lastColon + 1);
  const mapped = parseIpv4(tail);
  if (mapped !== null) return isPrivateIpv4(mapped);

  if (host === "::" || host === "::1") return true;
  const first = host.split(":")[0] ?? "";
  if (/^f[cd][0-9a-f]{0,2}$/.test(first)) return true; // fc00::/7 unique local
  if (/^fe[89ab][0-9a-f]?$/.test(first)) return true; // fe80::/10 link-local
  return false;
}

/** Builds the SSRF guard passed to {@link createImageResolver}. */
function createPrivateNetworkGuard(): RemoteGuard {
  return async (url: URL): Promise<ImageBytesResult | null> => {
    const hostname = bareHostname(url.hostname);
    const blocked = bytesFailure(
      "private-network-blocked",
      `host "${hostname}" resolves into private address space; refusing to fetch it`,
    );
    if (isPrivateAddress(hostname)) return blocked;
    if (parseIpv4(hostname) !== null || hostname.includes(":")) return null; // literal, already checked

    try {
      const addresses = await lookup(hostname, { all: true });
      if (addresses.some((entry) => isPrivateAddress(entry.address))) return blocked;
    } catch (error: unknown) {
      return bytesFailure("network-error", `DNS lookup failed: ${describeError(error)}`);
    }
    return null;
  };
}

/** Builds the filesystem loader passed to {@link createImageResolver}. */
function createFilesystemLoader(
  baseDir: string,
  allowOutside: boolean,
): (request: ImageLoadRequest) => Promise<ImageBytesResult | null> {
  const root = resolvePath(baseDir);

  return async ({ src, url, policy }: ImageLoadRequest): Promise<ImageBytesResult | null> => {
    let target: string;
    if (url !== null && url.protocol === "file:") {
      try {
        target = fileURLToPath(url);
      } catch (error: unknown) {
        return bytesFailure("invalid-url", `bad file: URL (${describeError(error)})`);
      }
    } else if (url !== null && !isWindowsPath(src)) {
      return null; // http(s), blob:, data: - not ours.
    } else {
      // A bare path. Strip a query/fragment the way a URL would, since
      // "./a.png?v=2" is common in markdown copied out of a site.
      target = src.replace(/[?#].*$/, "");
    }

    if (target === "") return bytesFailure("not-found", "empty path");

    const absolute = isAbsolute(target) ? resolvePath(target) : resolvePath(root, target);

    let real: string;
    try {
      real = await realpath(absolute);
    } catch (error: unknown) {
      return bytesFailure("not-found", `cannot read ${absolute} (${describeError(error)})`);
    }

    if (!allowOutside) {
      // Compare realpaths so a symlink cannot point out of the tree. `root`
      // itself is realpath'd lazily here rather than at construction time so a
      // missing baseDir is not a constructor error.
      let realRoot: string;
      try {
        realRoot = await realpath(root);
      } catch {
        realRoot = root;
      }
      if (
        real !== realRoot &&
        !real.startsWith(realRoot.endsWith(sep) ? realRoot : realRoot + sep)
      ) {
        return bytesFailure(
          "path-not-allowed",
          `${real} is outside the resolver's base directory (${realRoot}); ` +
            `pass { allowOutsideBaseDir: true } if that is intended`,
        );
      }
    }

    try {
      const info = await stat(real);
      if (!info.isFile()) return bytesFailure("not-found", `${real} is not a regular file`);
      if (info.size > policy.maxBytes) {
        return bytesFailure(
          "too-large",
          `${real} is ${info.size} bytes, over the ${policy.maxBytes}-byte cap`,
        );
      }
      if (info.size === 0) return bytesFailure("empty", `${real} is empty`);
      const buffer = await readFile(real);
      // Copy out of Node's Buffer: fs hands back a view into a pooled
      // allocation, and these bytes travel into a zip writer.
      return { ok: true, bytes: new Uint8Array(buffer), mediaType: null };
    } catch (error: unknown) {
      return bytesFailure("not-found", `cannot read ${real} (${describeError(error)})`);
    }
  };
}

/**
 * Builds an {@link ImageResolver} for Node.
 *
 * Same defaults as {@link createImageResolver} - **no network unless
 * `allowRemote: true`** - plus local files under `baseDir`.
 *
 * ```ts
 * const resolver = createNodeImageResolver({ baseDir: dirname(inputPath) });
 * const { images } = await resolveDocumentImages(doc, resolver);
 * ```
 *
 * @param options - See {@link NodeImageResolverOptions}.
 * @returns A resolver whose `resolve()` never throws and never rejects.
 */
export function createNodeImageResolver(options: NodeImageResolverOptions = {}): ImageResolver {
  const {
    baseDir,
    allowOutsideBaseDir = false,
    allowFilesystem = true,
    blockPrivateNetwork = true,
    ...rest
  } = options;

  const loaders = allowFilesystem
    ? [createFilesystemLoader(baseDir ?? process.cwd(), allowOutsideBaseDir)]
    : [];

  return createImageResolver({
    ...rest,
    loaders,
    ...(blockPrivateNetwork ? { guardRemote: createPrivateNetworkGuard() } : {}),
  });
}
