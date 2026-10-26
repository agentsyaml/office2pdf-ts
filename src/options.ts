import type { NativeConvertOptions } from "./native.js";

export interface ConvertOptions {
  pageRange?: string | null;
  page_range?: string | null;
  sheetNames?: string[] | null;
  sheet_names?: string[] | null;
  sheetFilter?: string[] | null;
  sheet_filter?: string[] | null;
  xlsxUiScript?: "hangul" | "latin" | null;
  xlsx_ui_script?: "hangul" | "latin" | null;
  slideRange?: string | null;
  slide_range?: string | null;
  includeHiddenSlides?: boolean | null;
  include_hidden_slides?: boolean | null;
  paperSize?: string | null;
  paper_size?: string | null;
  landscape?: boolean | null;
  tagged?: boolean | null;
  pdfUa?: boolean | null;
  pdf_ua?: boolean | null;
  fontPaths?: string[] | null;
  font_paths?: string[] | null;
  lastResortFontFamily?: string | null;
  last_resort_font_family?: string | null;
  pdfStandard?: string | null;
  pdf_standard?: string | null;
  includeWarnings?: boolean;
  include_warnings?: boolean;
  memoryLimitMb?: number | null;
  memory_limit_mb?: number | null;
  streamingChunkSize?: number | null;
  streaming_chunk_size?: number | null;
  streaming?: boolean;
}

function resolveAlias<T>(
  camelValue: T | undefined,
  snakeValue: T | undefined,
  field: string,
): T | undefined {
  const hasCamelValue = camelValue !== undefined && camelValue !== null;
  const hasSnakeValue = snakeValue !== undefined && snakeValue !== null;

  if (hasCamelValue && hasSnakeValue && camelValue !== snakeValue) {
    throw new TypeError(
      `${field} is ambiguous; provide either camelCase or snake_case, not both`,
    );
  }

  return camelValue ?? snakeValue;
}

function arraysEqual(
  left: string[] | null | undefined,
  right: string[] | null | undefined,
): boolean {
  return (
    left === right ||
    (!!left &&
      !!right &&
      left.length === right.length &&
      left.every((entry, index) => entry === right[index]))
  );
}

function normalizePositiveSafeInteger(
  value: number | null | undefined,
  fieldName: string,
): number | null | undefined {
  if (value === undefined || value === null) {
    return value;
  }

  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new TypeError(`${fieldName} must be a positive integer`);
  }

  return value;
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

function normalizeNonEmptyString(
  value: string | null | undefined,
  fieldName: string,
): string | null | undefined {
  if (value === undefined || value === null) {
    return value;
  }

  if (typeof value !== "string" || value.trim() === "") {
    throw new TypeError(`${fieldName} must be a non-empty string`);
  }

  return value;
}

function normalizeXlsxUiScript(
  value: string | null | undefined,
  fieldName: string,
): "hangul" | "latin" | null | undefined {
  if (value === undefined || value === null) {
    return value;
  }

  const normalized = String(value).trim().toLowerCase();
  if (normalized !== "hangul" && normalized !== "latin") {
    throw new TypeError(`${fieldName} must be "hangul" or "latin"`);
  }

  return normalized;
}

export function normalizeOptions(
  options: ConvertOptions = {},
): NativeConvertOptions {
  const pageRange = resolveAlias(
    options.pageRange,
    options.page_range,
    "pageRange/page_range",
  );
  if (pageRange !== undefined && pageRange !== null) {
    throw new Error("page_range is not supported by upstream office2pdf 0.9");
  }

  const memoryLimitMb = resolveAlias(
    options.memoryLimitMb,
    options.memory_limit_mb,
    "memoryLimitMb/memory_limit_mb",
  );
  if (memoryLimitMb !== undefined && memoryLimitMb !== null) {
    throw new Error(
      "memory_limit_mb is not supported by upstream office2pdf 0.9",
    );
  }

  const officialSheetNames = resolveAlias(
    normalizeOptionStringArray(options.sheetNames, "sheetNames"),
    normalizeOptionStringArray(options.sheet_names, "sheet_names"),
    "sheetNames/sheet_names",
  );
  const compatSheetNames = resolveAlias(
    normalizeOptionStringArray(options.sheetFilter, "sheetFilter"),
    normalizeOptionStringArray(options.sheet_filter, "sheet_filter"),
    "sheetFilter/sheet_filter",
  );
  const hasOfficialSheetNames =
    officialSheetNames !== undefined && officialSheetNames !== null;
  const hasCompatSheetNames =
    compatSheetNames !== undefined && compatSheetNames !== null;
  if (
    hasOfficialSheetNames &&
    hasCompatSheetNames &&
    !arraysEqual(officialSheetNames, compatSheetNames)
  ) {
    throw new TypeError(
      "sheetNames/sheet_names conflicts with sheetFilter/sheet_filter; provide matching values or use one alias family",
    );
  }

  const slideRange = resolveAlias(
    options.slideRange,
    options.slide_range,
    "slideRange/slide_range",
  );
  const xlsxUiScript = resolveAlias(
    normalizeXlsxUiScript(options.xlsxUiScript, "xlsxUiScript"),
    normalizeXlsxUiScript(options.xlsx_ui_script, "xlsx_ui_script"),
    "xlsxUiScript/xlsx_ui_script",
  );
  const includeHiddenSlides = resolveAlias(
    options.includeHiddenSlides,
    options.include_hidden_slides,
    "includeHiddenSlides/include_hidden_slides",
  );
  const paperSize = resolveAlias(
    options.paperSize,
    options.paper_size,
    "paperSize/paper_size",
  );
  const officialFontPaths = normalizeOptionStringArray(
    options.fontPaths,
    "fontPaths",
  );
  const compatFontPaths = normalizeOptionStringArray(
    options.font_paths,
    "font_paths",
  );
  const hasOfficialFontPaths =
    officialFontPaths !== undefined && officialFontPaths !== null;
  const hasCompatFontPaths =
    compatFontPaths !== undefined && compatFontPaths !== null;
  if (
    hasOfficialFontPaths &&
    hasCompatFontPaths &&
    !arraysEqual(officialFontPaths, compatFontPaths)
  ) {
    throw new TypeError(
      "fontPaths/font_paths conflicts; provide matching values or use one alias family",
    );
  }
  const fontPaths = hasOfficialFontPaths ? officialFontPaths : compatFontPaths;
  const lastResortFontFamily = resolveAlias(
    normalizeNonEmptyString(
      options.lastResortFontFamily,
      "lastResortFontFamily",
    ),
    normalizeNonEmptyString(
      options.last_resort_font_family,
      "last_resort_font_family",
    ),
    "lastResortFontFamily/last_resort_font_family",
  );
  const pdfStandard = resolveAlias(
    options.pdfStandard,
    options.pdf_standard,
    "pdfStandard/pdf_standard",
  );
  const landscape = resolveAlias(options.landscape, undefined, "landscape");
  const tagged = resolveAlias(options.tagged, undefined, "tagged");
  const pdfUa = resolveAlias(options.pdfUa, options.pdf_ua, "pdfUa/pdf_ua");
  const streamingChunkSize = normalizePositiveSafeInteger(
    resolveAlias(
      options.streamingChunkSize,
      options.streaming_chunk_size,
      "streamingChunkSize/streaming_chunk_size",
    ),
    "streamingChunkSize",
  );
  const includeWarnings = resolveAlias(
    options.includeWarnings,
    options.include_warnings,
    "includeWarnings",
  );
  const sheetNames = hasOfficialSheetNames
    ? officialSheetNames
    : compatSheetNames;

  return {
    sheet_names: sheetNames ?? null,
    xlsx_ui_script: xlsxUiScript ?? null,
    slide_range: slideRange ?? null,
    include_hidden_slides: includeHiddenSlides ?? null,
    paper_size: paperSize ?? null,
    landscape: landscape ?? null,
    tagged: tagged ?? null,
    pdf_ua: pdfUa ?? null,
    font_paths: fontPaths ?? [],
    last_resort_font_family: lastResortFontFamily ?? null,
    pdf_standard: pdfStandard ?? null,
    include_warnings: includeWarnings === undefined ? true : includeWarnings,
    streaming_chunk_size: streamingChunkSize ?? null,
    streaming: options.streaming ?? false,
  };
}
