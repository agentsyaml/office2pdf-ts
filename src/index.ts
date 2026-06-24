import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { extname } from "node:path";
import { getNativeConverter } from "./native.js";
import { normalizeOptions, type ConvertOptions } from "./options.js";

export type { ConvertOptions } from "./options.js";

export const __version__ = "0.1.1" as const;

export type Format = "docx" | "pptx" | "xlsx";

export const Format = {
  DOCX: "docx",
  PPTX: "pptx",
  XLSX: "xlsx",
  docx: "docx",
  pptx: "pptx",
  xlsx: "xlsx",
} as const;

export interface ConversionMetrics {
  parseDurationMs: number;
  codegenDurationMs: number;
  compileDurationMs: number;
  totalDurationMs: number;
  inputSizeBytes: number;
  outputSizeBytes: number;
  pageCount: number;
}

export interface ConversionResult {
  pdf: Uint8Array;
  warnings: string[];
  metrics: ConversionMetrics | null;
}

type ByteLike = Uint8Array | ArrayBuffer | Buffer;
type UnknownPathLike = string | URL;

function isByteLike(data: unknown): data is ByteLike {
  return data instanceof Uint8Array || data instanceof ArrayBuffer;
}

function coerceBytes(data: ByteLike): Uint8Array {
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  return data;
}

function normalizeFormat(value: Format | string): Format {
  const normalized = String(value).trim().replace(/^\./, "").toLowerCase();

  switch (normalized) {
    case Format.docx:
      return Format.docx;
    case Format.pptx:
      return Format.pptx;
    case Format.xlsx:
      return Format.xlsx;
    default:
      throw new Error("format must be one of: docx, pptx, xlsx");
  }
}

function normalizePathLike(path: UnknownPathLike): string {
  if (typeof path === "string") {
    return path;
  }

  if (path instanceof URL) {
    return path.protocol === "file:" ? fileURLToPath(path) : path.pathname;
  }

  throw new TypeError("path must be a string or URL-like path");
}

export function inferFormat(path: UnknownPathLike): Format {
  const pathString = normalizePathLike(path);
  const ext = extname(pathString);

  if (!ext) {
    throw new Error("path has no extension; expected .docx, .pptx, or .xlsx");
  }

  return normalizeFormat(ext);
}

export async function convertBytes(
  data: unknown,
  format: Format | string,
  options: ConvertOptions = {},
): Promise<ConversionResult> {
  if (!isByteLike(data)) {
    throw new TypeError("data must be bytes-like");
  }

  const bytes = coerceBytes(data);
  if (bytes.byteLength === 0) {
    throw new Error("data must not be empty");
  }

  const resolvedFormat = normalizeFormat(format);
  const normalizedOptions = normalizeOptions(options);
  const native = await getNativeConverter();

  const raw = await native(bytes, resolvedFormat, normalizedOptions);
  if (
    !raw ||
    typeof raw !== "object" ||
    !("pdf" in raw) ||
    !raw.pdf ||
    !(raw.pdf instanceof Uint8Array)
  ) {
    throw new TypeError("native conversion result did not contain PDF bytes");
  }

  const warningsSource = raw.warnings;
  if (
    !Array.isArray(warningsSource) ||
    !warningsSource.every((item) => typeof item === "string")
  ) {
    throw new TypeError(
      "native conversion result warnings must be a sequence of strings",
    );
  }

  return {
    pdf: raw.pdf,
    warnings: [...warningsSource],
    metrics:
      raw.metrics === undefined || raw.metrics === null ? null : raw.metrics,
  };
}

export async function convertPath(
  path: UnknownPathLike,
  options: ConvertOptions = {},
): Promise<ConversionResult> {
  const pathString = normalizePathLike(path);
  const inputFormat = inferFormat(pathString);
  const bytes = await readFile(pathString);

  return convertBytes(bytes, inputFormat, options);
}

export async function convertToPdf(
  data: unknown,
  format: Format | string,
): Promise<Uint8Array> {
  return (await convertBytes(data, format)).pdf;
}

export async function convertDocxToPdf(data: unknown): Promise<Uint8Array> {
  return convertToPdf(data, Format.docx);
}

export async function convertPptxToPdf(data: unknown): Promise<Uint8Array> {
  return convertToPdf(data, Format.pptx);
}

export async function convertXlsxToPdf(data: unknown): Promise<Uint8Array> {
  return convertToPdf(data, Format.xlsx);
}

export { Format as Formats };
