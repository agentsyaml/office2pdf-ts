import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  __version__,
  Format,
  convertDocxToPdf,
  convertBytes,
  convertPptxToPdf,
  convertPath,
  convertToPdf,
  convertXlsxToPdf,
  inferFormat,
  type ConversionResult,
} from "../src/index.js";
import { setNativeConverter } from "../src/native.js";

const nativeMock = vi.fn(async () => {
  const result: ConversionResult = {
    pdf: new Uint8Array([0x01, 0x02, 0x03]),
    warnings: ["warn"],
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
});

describe("public API", () => {
  it("exports __version__ and format enum", () => {
    expect(__version__).toBe("0.1.1");
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
    const unsupportedFormat: string = "txt";
    await expect(
      convertBytes(new Uint8Array([1, 2]), unsupportedFormat),
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

  it.each([
    [
      "convertToPdf",
      (input: Uint8Array) => convertToPdf(input, Format.docx),
      Format.docx,
    ],
    [
      "convertDocxToPdf",
      (input: Uint8Array) => convertDocxToPdf(input),
      Format.docx,
    ],
    [
      "convertPptxToPdf",
      (input: Uint8Array) => convertPptxToPdf(input),
      Format.pptx,
    ],
    [
      "convertXlsxToPdf",
      (input: Uint8Array) => convertXlsxToPdf(input),
      Format.xlsx,
    ],
  ])("%s returns PDF bytes only and uses %s", async (_, invoke, format) => {
    const pdf = await invoke(new Uint8Array([9, 8, 7]));

    expect(pdf).toEqual(new Uint8Array([1, 2, 3]));
    expect(nativeMock).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      format,
      expect.any(Object),
    );
  });

  it("normalizes upstream option aliases", async () => {
    await convertBytes(new Uint8Array([1, 2]), Format.docx, {
      sheetNames: ["One", "Two"],
      tagged: true,
      pdfUa: true,
      streamingChunkSize: 4096,
      slideRange: "1-3",
      paperSize: "A4",
      fontPaths: ["/tmp/font.ttf"],
      pdfStandard: "PdfA2b",
      includeWarnings: false,
      streaming: true,
    });

    expect(nativeMock).toHaveBeenCalledTimes(1);
    expect(nativeMock).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      Format.docx,
      expect.objectContaining({
        sheet_names: ["One", "Two"],
        tagged: true,
        pdf_ua: true,
        streaming_chunk_size: 4096,
        slide_range: "1-3",
        paper_size: "A4",
        font_paths: ["/tmp/font.ttf"],
        pdf_standard: "PdfA2b",
        include_warnings: false,
        streaming: true,
      }),
    );
  });

  it("normalizes snake_case aliases and sheetFilter compatibility", async () => {
    await convertBytes(new Uint8Array([1, 2]), Format.docx, {
      sheet_filter: ["Legacy"],
      pdf_ua: true,
      streaming_chunk_size: 2048,
    });

    expect(nativeMock).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      Format.docx,
      expect.objectContaining({
        sheet_names: ["Legacy"],
        pdf_ua: true,
        streaming_chunk_size: 2048,
      }),
    );
  });

  it.each([
    [
      "throws the ambiguity-style error for conflicting sheet aliases",
      { sheetNames: ["One"], sheet_names: ["Two"] },
      "sheetNames/sheet_names is ambiguous; provide either camelCase or snake_case, not both",
    ],
    [
      "throws the ambiguity-style error for conflicting font-path aliases",
      { fontPaths: ["/tmp/font-a.ttf"], font_paths: ["/tmp/font-b.ttf"] },
      "fontPaths/font_paths conflicts; provide matching values or use one alias family",
    ],
    [
      "rejects invalid streamingChunkSize before native conversion",
      { streamingChunkSize: 0 },
      "streamingChunkSize must be a positive integer",
    ],
  ])("%s", async (_, options, message) => {
    await expect(
      convertBytes(new Uint8Array([1, 2]), Format.docx, options),
    ).rejects.toThrow(message);
    expect(nativeMock).not.toHaveBeenCalled();
  });

  it("normalizes font path aliases by value when both provided", async () => {
    await convertBytes(new Uint8Array([1, 2]), Format.docx, {
      fontPaths: ["/tmp/font-a.ttf", "/tmp/font-b.ttf"],
      font_paths: ["/tmp/font-a.ttf", "/tmp/font-b.ttf"],
    });

    expect(nativeMock).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      Format.docx,
      expect.objectContaining({
        font_paths: ["/tmp/font-a.ttf", "/tmp/font-b.ttf"],
      }),
    );
  });

  it("convertPath validates extension before reading", async () => {
    await expect(convertPath("/tmp/readme.txt")).rejects.toThrow(
      "format must be one of: docx, pptx, xlsx",
    );

    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.docx");
      await fs.writeFile(inputPath, Buffer.from([1, 2, 3]));
      const result = await convertPath(inputPath);
      expect(result.pdf).toEqual(new Uint8Array([1, 2, 3]));
    });
  });
});
