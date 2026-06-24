import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { run } from "../src/cli.js";
import { __version__, Format, type ConversionResult } from "../src/index.js";
import { setNativeConverter } from "../src/native.js";

const nativeMock = vi.fn(async () => {
  const result: ConversionResult = {
    pdf: new Uint8Array([0x01, 0x02, 0x03]),
    warnings: [],
    metrics: null,
  };

  return result;
});

async function withWorkspace(
  test: (workspace: string) => Promise<void>,
): Promise<void> {
  const workspace = await fs.mkdtemp(join(tmpdir(), "office2pdf-ts-"));
  try {
    await test(workspace);
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
}

beforeEach(() => {
  vi.clearAllMocks();
  setNativeConverter(nativeMock);
  process.exitCode = 0;
});

afterEach(() => {
  setNativeConverter(null);
  vi.restoreAllMocks();
  process.exitCode = 0;
});

describe("CLI", () => {
  it("prints version", async () => {
    const stdoutWrite = vi
      .spyOn(process.stdout, "write")
      .mockImplementation(() => true);

    await run(["--version"]);

    expect(stdoutWrite).toHaveBeenCalledWith(`${__version__}\n`);
  });

  it("prints help for --help", async () => {
    const stdoutWrite = vi
      .spyOn(process.stdout, "write")
      .mockImplementation(() => true);

    await run(["--help"]);

    expect(stdoutWrite).toHaveBeenCalledWith(expect.stringContaining("Usage:"));
  });

  it("writes output PDF on success", async () => {
    const stderrWrite = vi
      .spyOn(process.stderr, "write")
      .mockImplementation(() => true);

    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.docx");
      const outputPath = join(workspace, "output.pdf");
      await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));

      await run([inputPath, outputPath]);

      expect(stderrWrite).not.toHaveBeenCalled();
      expect(process.exitCode).toBe(0);
      expect(await fs.readFile(outputPath)).toEqual(Buffer.from([1, 2, 3]));
    });
  });

  it("writes error on invalid call", async () => {
    const stderrWrite = vi
      .spyOn(process.stderr, "write")
      .mockImplementation(() => true);

    await run(["input.bin"]);

    expect(stderrWrite).toHaveBeenCalledWith(
      expect.stringContaining("expected input path and output path"),
    );
    expect(process.exitCode).toBe(1);
  });

  it("rejects ambiguous output target flags", async () => {
    const stderrWrite = vi
      .spyOn(process.stderr, "write")
      .mockImplementation(() => true);

    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.docx");
      await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));

      await run([
        inputPath,
        join(workspace, "output.pdf"),
        "--outdir",
        join(workspace, "out"),
      ]);

      expect(stderrWrite).toHaveBeenCalledWith(
        expect.stringContaining("output path is ambiguous"),
      );
      expect(process.exitCode).toBe(1);
    });
  });

  it.each([
    {
      name: "accepts -o as an output flag",
      args: (inputPath: string, workspace: string) => [
        inputPath,
        "-o",
        join(workspace, "output.pdf"),
      ],
      outputPath: (workspace: string) => join(workspace, "output.pdf"),
    },
    {
      name: "accepts --outdir as an output directory flag",
      args: (inputPath: string, workspace: string) => [
        inputPath,
        "--outdir",
        join(workspace, "out"),
      ],
      outputPath: (workspace: string) => join(workspace, "out", "input.pdf"),
    },
  ])("$name", async ({ args, outputPath }) => {
    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.docx");
      await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));

      await run(args(inputPath, workspace));

      expect(await fs.readFile(outputPath(workspace))).toEqual(
        Buffer.from([1, 2, 3]),
      );
    });
  });

  it("maps CLI paper, sheet, slide, pdf-a, landscape, and font flags", async () => {
    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.xlsx");
      const outputPath = join(workspace, "output.pdf");
      await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));

      await run([
        "--paper",
        "A4",
        "--landscape",
        "--pdf-a",
        "--sheets",
        "Sheet1,Sheet2",
        "--slides",
        "1-3",
        "--font-path",
        "/tmp/font-a.ttf",
        "--font-path",
        "/tmp/font-b.ttf",
        inputPath,
        outputPath,
      ]);

      expect(nativeMock).toHaveBeenCalledWith(
        expect.any(Uint8Array),
        Format.xlsx,
        expect.objectContaining({
          paper_size: "A4",
          landscape: true,
          pdf_standard: "pdf/a-2b",
          sheet_names: ["Sheet1", "Sheet2"],
          slide_range: "1-3",
          font_paths: ["/tmp/font-a.ttf", "/tmp/font-b.ttf"],
        }),
      );
    });
  });
});
