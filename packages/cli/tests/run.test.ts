/**
 * `run()` against a fake process.
 *
 * Two things cannot be checked by spawning the binary. A child process's stdout
 * is a pipe, never a terminal, so the TTY guard would never fire; and a test
 * that allocated a pty to prove it would be testing the pty. `run()` takes its
 * three streams as an argument precisely so that both can be stated directly.
 *
 * Everything else about the CLI is exercised through the real binary — see
 * `e2e.test.ts`.
 */

import { Readable, Writable } from "node:stream";

import { describe, expect, it } from "vitest";

import { EXIT } from "../src/errors.js";
import type { Io } from "../src/io.js";
import { run } from "../src/run.js";

/** A stream that keeps what was written to it. */
class Sink extends Writable {
  readonly chunks: Buffer[] = [];
  /** Pretends to be a terminal, which is the whole point of this file. */
  isTTY = false;

  override _write(
    chunk: Buffer | string,
    encoding: BufferEncoding,
    done: (error?: Error | null) => void,
  ): void {
    this.chunks.push(typeof chunk === "string" ? Buffer.from(chunk, encoding) : chunk);
    done();
  }

  get bytes(): Buffer {
    return Buffer.concat(this.chunks);
  }

  get text(): string {
    return this.bytes.toString("utf8");
  }
}

/** A fake process, with a terminal wherever you ask for one. */
function fakeIo(options: { stdin?: string; stdoutIsTty?: boolean; stdinIsTty?: boolean } = {}): {
  io: Io;
  stdout: Sink;
  stderr: Sink;
} {
  const stdout = new Sink();
  stdout.isTTY = options.stdoutIsTty ?? false;
  const stderr = new Sink();
  return {
    io: {
      stdin: Readable.from([options.stdin ?? ""]),
      stdout,
      stderr,
      stdinIsTty: options.stdinIsTty ?? false,
      cwd: process.cwd(),
    },
    stdout,
    stderr,
  };
}

describe("refusing to write binary to a terminal", () => {
  it("exits EX_USAGE (64) with an actionable message", async () => {
    const { io, stdout, stderr } = fakeIo({ stdin: "# hello\n", stdoutIsTty: true });

    const code = await run([], io);

    expect(code).toBe(EXIT.usage);
    expect(stderr.text).toContain("refusing to write a .docx to the terminal");
    expect(stderr.text).toContain("downword in.md > out.docx");
    expect(stdout.bytes.length).toBe(0);
  });

  it("checks before converting, not after", async () => {
    // A document big enough that converting it would be visible if it happened.
    const markdown = `# Heading\n\n${"paragraph text ".repeat(5_000)}\n`;
    const { io, stdout } = fakeIo({ stdin: markdown, stdoutIsTty: true });

    const started = Date.now();
    const code = await run([], io);

    expect(code).toBe(EXIT.usage);
    expect(stdout.bytes.length).toBe(0);
    // Not a benchmark: a real conversion of this input is far slower than this.
    expect(Date.now() - started).toBeLessThan(1_000);
  });

  it("writes the document happily when stdout is a pipe", async () => {
    const { io, stdout } = fakeIo({ stdin: "# piped\n" });

    const code = await run([], io);

    expect(code).toBe(EXIT.ok);
    expect(stdout.bytes.subarray(0, 2).toString("latin1")).toBe("PK");
  });

  it("prints the help when there is neither an argument nor a pipe", async () => {
    const { io, stdout } = fakeIo({ stdinIsTty: true, stdoutIsTty: true });

    const code = await run([], io);

    expect(code).toBe(EXIT.ok);
    expect(stdout.text).toContain("Usage:");
  });
});

describe("combinations that cannot be honoured", () => {
  it("rejects --outdir with stdin, which has no name to derive one from", async () => {
    const { io, stderr } = fakeIo({ stdin: "# hi\n" });

    const code = await run(["--outdir", "build"], io);

    expect(code).toBe(EXIT.usage);
    expect(stderr.text).toContain("--outdir needs input files");
  });
});

describe("run() is total", () => {
  it("returns a code rather than throwing, for every failure it knows", async () => {
    const cases: readonly (readonly string[])[] = [
      ["--nope"],
      ["--theme", "nope", "x.md"],
      ["does-not-exist.md"],
      ["x.md", "-o", "no/such/dir/x.docx"],
    ];

    for (const argv of cases) {
      const { io } = fakeIo({ stdinIsTty: false });
      await expect(run(argv, io)).resolves.not.toBe(EXIT.ok);
    }
  });
});
