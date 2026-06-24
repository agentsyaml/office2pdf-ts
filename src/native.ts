import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import type { ConversionResult } from "./index.js";

export interface NativeConvertOptions {
  sheet_names: string[] | null;
  slide_range: string | null;
  paper_size: string | null;
  landscape: boolean | null;
  tagged: boolean | null;
  pdf_ua: boolean | null;
  font_paths: string[];
  pdf_standard: string | null;
  include_warnings: boolean;
  streaming_chunk_size: number | null;
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

let override: NativeConverter | null = null;
let cached: NativeConverter | null = null;
const requireNative = createRequire(import.meta.url);

type UnknownRecord = Record<string, unknown>;

function isRecord(value: unknown): value is UnknownRecord {
  return value !== null && typeof value === "object";
}

function resolveImportPaths(): string[] {
  return [
    new URL("./wasm/office2pdf_wasm.cjs", import.meta.url).toString(),
    new URL("./wasm/office2pdf_wasm.js", import.meta.url).toString(),
    new URL("./wasm/pkg/office2pdf_wasm.js", import.meta.url).toString(),
    new URL("../wasm/pkg/office2pdf_wasm.js", import.meta.url).toString(),
  ];
}

function extractConverter(module: unknown): NativeConverter {
  const normalized = normalizeModuleExports(module);
  const converter = normalized.convert_bytes;

  if (typeof converter === "function") {
    return converter as NativeConverter;
  }

  const maybeDefault = normalized.default;
  if (maybeDefault && typeof maybeDefault === "object") {
    const normalizedDefault = normalizeModuleExports(maybeDefault);
    const defaultConverter = normalizedDefault.convert_bytes;
    if (typeof defaultConverter === "function") {
      return defaultConverter as NativeConverter;
    }
  }

  throw new Error("generated wasm module does not expose convert_bytes");
}

function normalizeModuleExports(module: unknown): UnknownRecord {
  if (!isRecord(module)) {
    throw new Error("invalid generated wasm module");
  }

  const direct = module as UnknownRecord;
  const nested = direct.default;

  if (isRecord(nested)) {
    return { ...nested, ...direct } as UnknownRecord;
  }

  return direct;
}

async function loadWasmModule(candidate: string): Promise<unknown> {
  const candidatePath = fileURLToPath(candidate);

  if (candidate.endsWith(".cjs") || candidate.endsWith(".js")) {
    try {
      return requireNative(candidatePath);
    } catch (requireError) {
      if (candidate.endsWith(".cjs")) {
        throw requireError;
      }
    }
  }

  return import(candidate);
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
      const module = await loadWasmModule(candidate);
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
