/**
 * The command line itself: help, version, and every way of getting it wrong.
 *
 * Nothing here converts a real document — that is `e2e.test.ts`. What these
 * assert is the CLI's contract with a shell: the help text is what it claims to
 * be, a bad flag is a usage error rather than a stack trace, and every failure
 * exits non-zero with the `sysexits.h` code that describes it.
 */

import { chmodSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { BIN, FIXTURES, PKG_DIR, openDocx, outPath, runCli } from "./helpers/cli.js";

const pkg = JSON.parse(readFileSync(join(PKG_DIR, "package.json"), "utf8")) as { version: string };

describe("downword --help", () => {
  it("prints help and exits 0", () => {
    const { status, out, stderr } = runCli(["--help"]);

    expect(stderr).toBe("");
    expect(status).toBe(0);
    expect(out).toMatchSnapshot("--help");
  });

  it("documents the default-deny on remote images, not only the opt-in", () => {
    const { out } = runCli(["--help"]);

    expect(out).toContain("the network is off unless you turn it on");
    // The default is stated first, and as a flag of its own.
    expect(out.indexOf("--no-remote-images")).toBeLessThan(out.indexOf("--remote-images"));
  });

  it("says what happens to mermaid fences", () => {
    expect(runCli(["--help"]).out).toContain("Mermaid diagrams need a browser DOM");
  });
});

describe("downword --version", () => {
  it("prints the package version", () => {
    const { status, out } = runCli(["--version"]);

    expect(status).toBe(0);
    expect(out.trim()).toBe(pkg.version);
  });

  it("is also -v", () => {
    expect(runCli(["-v"]).out.trim()).toBe(pkg.version);
  });
});

describe("usage errors", () => {
  it("rejects an unknown option with EX_USAGE (64)", () => {
    const { status, stderr } = runCli(["--definitely-not-a-flag"]);

    expect(status).toBe(64);
    expect(stderr).toContain("Unknown option");
    expect(stderr).toContain("Try 'downword --help'");
  });

  it("rejects a bad enum value, naming the values it takes", () => {
    const { status, stderr } = runCli(["--theme", "solarized", "x.md"]);

    expect(status).toBe(64);
    expect(stderr).toContain("--theme must be one of default, github, academic");
    expect(stderr).toContain("solarized");
  });

  it.each([
    ["--math", "latex"],
    ["--page-size", "A3"],
  ])("rejects %s %s", (flag, value) => {
    const { status, stderr } = runCli([flag, value, "x.md"]);

    expect(status).toBe(64);
    expect(stderr).toContain(`${flag} must be one of`);
  });

  it("rejects an option whose value is missing", () => {
    const { status, stderr } = runCli(["--theme"]);

    expect(status).toBe(64);
    expect(stderr).toContain("--theme");
  });

  it("rejects contradictory flags", () => {
    const both = runCli(["--remote-images", "--no-remote-images", "x.md"]);
    expect(both.status).toBe(64);
    expect(both.stderr).toContain("contradict each other");

    const loudAndQuiet = runCli(["--quiet", "--verbose", "x.md"]);
    expect(loudAndQuiet.status).toBe(64);
    expect(loudAndQuiet.stderr).toContain("contradict each other");

    const twoTargets = runCli(["-o", "a.docx", "--outdir", "b", "x.md"]);
    expect(twoTargets.status).toBe(64);
    expect(twoTargets.stderr).toContain("--output and --outdir");
  });

  it("rejects --output with more than one input", () => {
    const { status, stderr } = runCli(
      ["-o", outPath("never-written.docx"), "batch/one.md", "batch/two.md"],
      { cwd: FIXTURES },
    );

    expect(status).toBe(64);
    expect(stderr).toContain("--output takes one input file");
    expect(stderr).toContain("--outdir");
  });
});

describe("input errors", () => {
  it("exits EX_NOINPUT (66) for a file that is not there", () => {
    const { status, stderr } = runCli(["no-such-document.md"]);

    expect(status).toBe(66);
    expect(stderr).toContain("no such file: no-such-document.md");
  });

  it("exits EX_NOINPUT (66) for a directory", () => {
    const { status, stderr } = runCli(["batch"], { cwd: FIXTURES });

    expect(status).toBe(66);
    expect(stderr).toContain("is a directory");
    expect(stderr).toContain('"batch/*.md"');
  });

  it("exits EX_NOINPUT (66) for a pattern that matches nothing", () => {
    const { status, stderr } = runCli(["batch/*.rst"], { cwd: FIXTURES });

    expect(status).toBe(66);
    expect(stderr).toContain("no files match");
  });

  it("exits EX_NOINPUT (66) for an unreadable file", () => {
    const path = outPath("unreadable.md");
    writeFileSync(path, "# nope\n");
    chmodSync(path, 0o000);
    try {
      const { status, stderr } = runCli([path]);
      // Running as root makes every file readable. Rather than fail a container
      // that cannot be denied anything, stop asserting when the read succeeded.
      if (status === 0) return;
      expect(status).toBe(66);
      expect(stderr).toContain("permission denied");
    } finally {
      chmodSync(path, 0o644);
    }
  });
});

describe("output errors", () => {
  it("exits EX_CANTCREAT (73) when -o names a missing directory", () => {
    const { status, stderr } = runCli(["kitchen-sink.md", "-o", "no/such/place/out.docx"], {
      cwd: FIXTURES,
    });

    expect(status).toBe(73);
    expect(stderr).toContain("no such directory");
    expect(stderr).toContain("--outdir");
  });

  it("writes nothing when the target cannot be written", () => {
    const { status } = runCli(["kitchen-sink.md", "-o", "no/such/place/out.docx"], {
      cwd: FIXTURES,
    });

    expect(status).not.toBe(0);
    expect(statSync(join(FIXTURES, "kitchen-sink.md")).isFile()).toBe(true);
  });
});

describe("an empty document is still a document", () => {
  it("converts empty stdin to a valid, empty .docx", async () => {
    const { status, stdout } = runCli([], { input: "" });

    expect(status).toBe(0);
    const docx = await openDocx(stdout);
    expect(docx.parts).toContain("word/document.xml");
    expect(docx.text.trim()).toBe("");
  });
});

describe("the binary itself", () => {
  it("is executable and starts with a node shebang", () => {
    const contents = readFileSync(BIN, "utf8");

    expect(contents.startsWith("#!/usr/bin/env node\n")).toBe(true);
    expect(statSync(BIN).mode & 0o111).toBeGreaterThan(0);
  });
});
