import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { extname } from "node:path";
import { getNativeConverter, type NativeConvertOptions } from "./native.js";

export const __version__ = "0.1.0" as const;

export type Format = "docx" | "pptx" | "xlsx";

export const Format = {
  DOCX: "docx",
  PPTX: "pptx",
  XLSX: "xlsx",
  docx: "docx",
  pptx: "pptx",
  xlsx: "xlsx",
} as const;

export interface ConvertOptions {
  pageRange?: string | null;
  page_range?: string | null;
  sheetFilter?: string[] | null;
  sheet_filter?: string[] | null;
  slideRange?: string | null;
  slide_range?: string | null;
  paperSize?: string | null;
  paper_size?: string | null;
  landscape?: boolean | null;
  fontPaths?: string[] | null;
  font_paths?: string[] | null;
  pdfStandard?: string | null;
  pdf_standard?: string | null;
  includeWarnings?: boolean;
  include_warnings?: boolean;
  memoryLimitMb?: number | null;
  memory_limit_mb?: number | null;
  streaming?: boolean;
}

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

const SUPPORTED_FORMATS = new Set<Format>(["docx", "pptx", "xlsx"]);

function isByteLike(data: unknown): data is ByteLike {
  return data instanceof Uint8Array || data instanceof ArrayBuffer;
}

function coerceBytes(data: ByteLike): Uint8Array {
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  return data;
}

function resolveAlias<T>(
  camelValue: T | undefined,
  snakeValue: T | undefined,
  field: string,
): T | undefined {
  if (
    camelValue !== undefined &&
    snakeValue !== undefined &&
    camelValue !== snakeValue
  ) {
    throw new TypeError(
      `${field} is ambiguous; provide either camelCase or snake_case, not both`,
    );
  }

  return camelValue ?? snakeValue;
}

function normalizeFormat(value: Format | string): Format {
  const normalized = String(value).trim().replace(/^\./, "").toLowerCase();

  if (!SUPPORTED_FORMATS.has(normalized as Format)) {
    throw new Error("format must be one of: docx, pptx, xlsx");
  }

  return normalized as Format;
}

function normalizePathLike(path: UnknownPathLike): string {
  if (typeof path === "string") {
    return path;
  }

  if (path instanceof URL) {
    if (path.protocol === "file:") {
      return fileURLToPath(path);
    }

    return path.pathname;
  }

  throw new TypeError("path must be a string or URL-like path");
}

function normalizeOptionStringArray(
  value: string[] | null | undefined,
  fieldName: string,
): string[] | null | undefined {
  if (value === undefined || value === null) {
    return value;
  }

  if (
    !Array.isArray(value) ||
    !value.every((entry) => typeof entry === "string")
  ) {
    throw new TypeError(`${fieldName} must be an array of strings`);
  }

  return value;
}

function normalizeOptions(options: ConvertOptions = {}): NativeConvertOptions {
  const pageRange = resolveAlias(
    options.pageRange,
    options.page_range,
    "pageRange/page_range",
  );
  if (pageRange !== undefined && pageRange !== null) {
    throw new Error("page_range is not supported by upstream office2pdf 0.6");
  }

  const memoryLimitMb = resolveAlias(
    options.memoryLimitMb,
    options.memory_limit_mb,
    "memoryLimitMb/memory_limit_mb",
  );
  if (memoryLimitMb !== undefined && memoryLimitMb !== null) {
    throw new Error(
      "memory_limit_mb is not supported by upstream office2pdf 0.6",
    );
  }

  const sheetFilter = normalizeOptionStringArray(
    resolveAlias(
      options.sheetFilter,
      options.sheet_filter,
      "sheetFilter/sheet_filter",
    ),
    "sheetFilter",
  );
  const slideRange = resolveAlias(
    options.slideRange,
    options.slide_range,
    "slideRange/slide_range",
  );
  const paperSize = resolveAlias(
    options.paperSize,
    options.paper_size,
    "paperSize/paper_size",
  );
  const fontPaths = normalizeOptionStringArray(
    resolveAlias(options.fontPaths, options.font_paths, "fontPaths/font_paths"),
    "fontPaths",
  );
  const pdfStandard = resolveAlias(
    options.pdfStandard,
    options.pdf_standard,
    "pdfStandard/pdf_standard",
  );
  const landscape = resolveAlias(options.landscape, undefined, "landscape");
  const includeWarnings = resolveAlias(
    options.includeWarnings,
    options.include_warnings,
    "includeWarnings",
  );
  const streaming = options.streaming;

  return {
    sheet_filter: sheetFilter ?? null,
    slide_range: slideRange ?? null,
    paper_size: paperSize ?? null,
    landscape: landscape ?? null,
    font_paths: fontPaths ?? [],
    pdf_standard: pdfStandard ?? null,
    include_warnings: includeWarnings === undefined ? true : includeWarnings,
    streaming: streaming ?? false,
  };
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

export { Format as Formats };
