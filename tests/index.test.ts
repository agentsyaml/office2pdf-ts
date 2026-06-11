import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  __version__,
  ConversionResult,
  Format,
  convertBytes,
  convertPath,
  inferFormat,
} from "../src/index.js";
import { setNativeConverter } from "../src/native.js";
import { run } from "../src/cli.js";

const nativeMock = vi.fn(async () => {
  const result: ConversionResult = {
    pdf: new Uint8Array([0x01, 0x02, 0x03]),
    warnings: ["warn"],
    metrics: null,
  };

  return result;
});

beforeEach(() => {
  vi.clearAllMocks();
  setNativeConverter(nativeMock);
  process.exitCode = 0;
});

afterEach(() => {
  setNativeConverter(null);
  vi.restoreAllMocks();
});

describe("public API", () => {
  it("exports __version__ and format enum", () => {
    expect(__version__).toBe("0.1.0");
    expect(Format.DOCX).toBe("docx");
    expect(Format.PPTX).toBe("pptx");
    expect(Format.XLSX).toBe("xlsx");
    expect(Format.docx).toBe("docx");
  });

  it("infers format from filename and URL", () => {
    expect(inferFormat("/tmp/sample.docx")).toBe("docx");
    expect(inferFormat(new URL("https://example.org/path/data.xlsx"))).toBe(
      "xlsx",
    );
    expect(() => inferFormat("/tmp/unknown.bin")).toThrow(
      "format must be one of: docx, pptx, xlsx",
    );
  });

  it("validates convertBytes input and rejects unsupported options", async () => {
    await expect(convertBytes({}, Format.docx)).rejects.toThrow(
      "data must be bytes-like",
    );
    await expect(convertBytes(new Uint8Array(0), Format.docx)).rejects.toThrow(
      "data must not be empty",
    );
    await expect(
      convertBytes(new Uint8Array([1, 2]), "txt" as string),
    ).rejects.toThrow("format must be one of: docx, pptx, xlsx");
    await expect(
      convertBytes(new Uint8Array([1, 2]), Format.docx, { pageRange: "1-2" }),
    ).rejects.toThrow("page_range is not supported by upstream office2pdf 0.6");
    await expect(
      convertBytes(new Uint8Array([1, 2]), Format.docx, {
        memory_limit_mb: 1024,
      }),
    ).rejects.toThrow(
      "memory_limit_mb is not supported by upstream office2pdf 0.6",
    );
  });

  it("normalizes options and passes through snake/camel aliases", async () => {
    await convertBytes(new Uint8Array([1, 2]), Format.docx, {
      sheetFilter: ["One", "Two"],
      slide_range: "1-3",
      paperSize: "A4",
      landscape: null,
      font_paths: ["/tmp/font.ttf"],
      pdf_standard: "PdfA2b",
      include_warnings: false,
      streaming: true,
    });

    expect(nativeMock).toHaveBeenCalledTimes(1);
    expect(nativeMock).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      Format.docx,
      expect.objectContaining({
        sheet_filter: ["One", "Two"],
        slide_range: "1-3",
        paper_size: "A4",
        landscape: null,
        font_paths: ["/tmp/font.ttf"],
        pdf_standard: "PdfA2b",
        include_warnings: false,
        streaming: true,
      }),
    );
  });

  it("convertPath validates extension before reading", async () => {
    await expect(convertPath("/tmp/readme.txt")).rejects.toThrow(
      "format must be one of: docx, pptx, xlsx",
    );

    const workspace = await fs.mkdtemp(join(tmpdir(), "office2pdf-ts-"));
    const inputPath = join(workspace, "input.docx");
    await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));
    const result = await convertPath(inputPath);
    expect(result.pdf).toEqual(new Uint8Array([1, 2, 3]));

    await fs.rm(workspace, { recursive: true, force: true });
  });
});

describe("CLI", () => {
  afterEach(() => {
    process.exitCode = 0;
  });

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
    const workspace = await fs.mkdtemp(join(tmpdir(), "office2pdf-ts-"));
    const inputPath = join(workspace, "input.docx");
    const outputPath = join(workspace, "output.pdf");
    await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));

    await run([inputPath, outputPath]);
    expect(stderrWrite).not.toHaveBeenCalled();
    expect(process.exitCode).toBe(0);

    const out = await fs.readFile(outputPath);
    expect(out).toEqual(Buffer.from([0x01, 0x02, 0x03]));

    await fs.rm(workspace, { recursive: true, force: true });
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
});
