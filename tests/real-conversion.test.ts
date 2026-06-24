import { describe, expect, it } from "vitest";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { TextDecoder } from "node:util";
import { Format, convertBytes, convertPath } from "../src/index.js";
import {
  createDocxBytes,
  createPptxBytes,
  createXlsxBytes,
} from "./helpers/office-fixtures.js";

const decoder = new TextDecoder("utf-8");

async function withWorkspace(
  test: (workspace: string) => Promise<void>,
): Promise<void> {
  const workspace = await fs.mkdtemp(join(tmpdir(), "office2pdf-ts-real-"));
  try {
    await test(workspace);
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
}

function expectPdf(bytes: Uint8Array): void {
  expect(bytes).toBeInstanceOf(Uint8Array);
  expect(bytes.length).toBeGreaterThan(0);
  expect(decoder.decode(bytes.slice(0, 5))).toBe("%PDF-");
}

describe("real conversion", () => {
  it("converts real DOCX fixture with convertBytes", async () => {
    const result = await convertBytes(createDocxBytes(), Format.docx);

    expectPdf(result.pdf);
    expect(result.warnings).toEqual(expect.any(Array));

    if (result.metrics !== null) {
      expect(result.metrics).toMatchObject({
        parseDurationMs: expect.any(Number),
        codegenDurationMs: expect.any(Number),
        compileDurationMs: expect.any(Number),
        totalDurationMs: expect.any(Number),
        inputSizeBytes: expect.any(Number),
        outputSizeBytes: expect.any(Number),
        pageCount: expect.any(Number),
      });
    }
  });

  it("converts real PPTX fixture with convertBytes", async () => {
    const result = await convertBytes(createPptxBytes(), Format.pptx);

    expectPdf(result.pdf);
    expect(result.warnings).toEqual(expect.any(Array));
  });

  it("converts real XLSX fixture with convertBytes", async () => {
    const result = await convertBytes(createXlsxBytes(), Format.xlsx);

    expectPdf(result.pdf);
    expect(result.warnings).toEqual(expect.any(Array));
  });

  it("accepts upstream-only options during real conversion", async () => {
    const result = await convertBytes(createXlsxBytes(), Format.xlsx, {
      tagged: true,
      pdfUa: false,
      streaming: true,
      streamingChunkSize: 1024,
    });

    expectPdf(result.pdf);
    expect(result.warnings).toEqual(expect.any(Array));
  });

  it("converts real DOCX file with convertPath", async () => {
    await withWorkspace(async (workspace) => {
      const inputPath = join(workspace, "input.docx");
      const outputPath = join(workspace, "output.pdf");
      await fs.writeFile(inputPath, Buffer.from(createDocxBytes()));

      const result = await convertPath(inputPath);
      await fs.writeFile(outputPath, result.pdf);
      const outputPdf = await fs.readFile(outputPath);

      expect(outputPdf.slice(0, 5).toString()).toBe("%PDF-");
    });
  });
});
