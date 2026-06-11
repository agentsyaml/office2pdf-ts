import type { ConversionResult } from "./index.js";

export interface NativeConvertOptions {
  sheet_filter: string[] | null;
  slide_range: string | null;
  paper_size: string | null;
  landscape: boolean | null;
  font_paths: string[];
  pdf_standard: string | null;
  include_warnings: boolean;
  streaming: boolean;
}

export type NativeConverter = (
  bytes: Uint8Array,
  format: string,
  options: NativeConvertOptions,
) => Promise<{
  pdf: Uint8Array;
  warnings: string[];
  metrics: ConversionResult["metrics"];
}>;

type NativeModule = {
  convert_bytes: NativeConverter;
};

let override: NativeConverter | null = null;
let cached: NativeConverter | null = null;

function resolveImportPaths(): string[] {
  return [
    new URL("./wasm/office2pdf_wasm.js", import.meta.url).toString(),
    new URL("./wasm/pkg/office2pdf_wasm.js", import.meta.url).toString(),
    new URL("../wasm/pkg/office2pdf_wasm.js", import.meta.url).toString(),
  ];
}

function extractConverter(module: unknown): NativeConverter {
  if (!module || typeof module !== "object") {
    throw new Error("invalid generated wasm module");
  }

  const converter = (module as Partial<NativeModule>).convert_bytes;
  if (typeof converter !== "function") {
    throw new Error("generated wasm module does not expose convert_bytes");
  }

  return converter;
}

export function setNativeConverter(converter: NativeConverter | null): void {
  override = converter;
  cached = converter;
}

export async function getNativeConverter(): Promise<NativeConverter> {
  if (override) {
    return override;
  }

  if (cached) {
    return cached;
  }

  const candidates = resolveImportPaths();

  let lastError: unknown = null;
  for (const candidate of candidates) {
    try {
      const module = (await import(candidate)) as unknown;
      cached = extractConverter(module);
      return cached;
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(
    `failed to load native wasm module from ${candidates.join(", ")} (${String(lastError)})`,
  );
}
