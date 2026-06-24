# @alexsun-top/office2pdf

Node-first TypeScript binding for [`office2pdf = "0.6"`](https://docs.rs/office2pdf/0.6.0/office2pdf/) via a generated WebAssembly bridge.

- No LibreOffice service
- No Chromium/Docker/browser-only wrappers
- No Python/CLI dependency

## Install

```bash
npm install @alexsun-top/office2pdf
```

## Usage

```ts
import {
  convertBytes,
  convertDocxToPdf,
  convertPath,
  convertToPdf,
  inferFormat,
  __version__,
  Format,
} from "@alexsun-top/office2pdf";
import { writeFile } from "node:fs/promises";

// infer extension before conversion
const format = inferFormat("/tmp/report.docx"); // "docx"

// convert local file
const result = await convertPath("/tmp/report.docx");
await writeFile("/tmp/report.pdf", result.pdf);

// convert bytes
const input = new Uint8Array([
  /* office file bytes */
]);
const { pdf, warnings, metrics } = await convertBytes(input, Format.DOCX, {
  includeWarnings: true,
  paperSize: "A4",
});

// upstream-compatible JS/WASM helper style: returns PDF bytes only
const helperPdf = await convertDocxToPdf(input);
const genericPdf = await convertToPdf(input, "docx");
```

### API

- `__version__: string`
- `Format: { DOCX: "docx"; PPTX: "pptx"; XLSX: "xlsx"; docx: "docx"; pptx: "pptx"; xlsx: "xlsx" }`
- `inferFormat(pathOrUrl: string | URL): "docx" | "pptx" | "xlsx"`
- `convertBytes(data, format, options?)`
- `convertPath(pathOrUrl, options?)`
- `convertToPdf(data, format)` -> `Promise<Uint8Array>`
- `convertDocxToPdf(data)` -> `Promise<Uint8Array>`
- `convertPptxToPdf(data)` -> `Promise<Uint8Array>`
- `convertXlsxToPdf(data)` -> `Promise<Uint8Array>`

`ConvertOptions` accepts both camelCase and snake_case aliases:

- `sheetNames` / `sheet_names`
- `sheetFilter` / `sheet_filter` compatibility alias for `sheet_names`
- `slideRange` / `slide_range`
- `paperSize` / `paper_size`
- `fontPaths` / `font_paths`
- `pdfStandard` / `pdf_standard`
- `includeWarnings` / `include_warnings`
- `pdfUa` / `pdf_ua`
- `streamingChunkSize` / `streaming_chunk_size`

It also accepts upstream boolean options `tagged` and `streaming`.

Rust-native upstream APIs map to this Node package as follows:

- `office2pdf::convert(path)` and `convert_with_options(path, options)` correspond to `convertPath(path, options?)`.
- `office2pdf::convert_bytes(data, format, options)` corresponds to `convertBytes(data, format, options?)`.
- `render_document` is not exposed because this package does not expose the upstream IR document model.

Unsupported by upstream `office2pdf 0.6`:

- `pageRange` / `page_range` -> throws
  - `page_range is not supported by upstream office2pdf 0.6`
- `memoryLimitMb` / `memory_limit_mb` -> throws
  - `memory_limit_mb is not supported by upstream office2pdf 0.6`

#### Result shape

```ts
{
  pdf: Uint8Array;
  warnings: string[];
  metrics: null | {
    parseDurationMs: number;
    codegenDurationMs: number;
    compileDurationMs: number;
    totalDurationMs: number;
    inputSizeBytes: number;
    outputSizeBytes: number;
    pageCount: number;
  };
}
```

## CLI

Build exports a `office2pdf` command:

```bash
office2pdf --help
office2pdf --version
office2pdf input.docx output.pdf
office2pdf input.docx -o output.pdf
office2pdf input.xlsx --outdir ./pdfs
office2pdf --paper A4 --landscape --pdf-a input.docx output.pdf
office2pdf --sheets Sheet1,Sheet2 input.xlsx output.pdf
office2pdf --slides 1-3 input.pptx output.pdf
office2pdf --font-path /path/font-a.ttf --font-path /path/font-b.ttf input.docx output.pdf
```

Supported flags: `-o` / `--output`, `--outdir`, `--paper`, `--landscape`, `--pdf-a`, `--sheets`, `--slides`, and repeatable or comma-separated `--font-path`.

## Scripts

- `npm run wasm:build` — build wasm artifacts to `wasm/pkg`
- `npm run build` — build wasm (release), copy publishable wasm files, and compile TypeScript
- `npm run typecheck` — TypeScript typecheck
- `npm test` — run Vitest
- `npm run check` — typecheck + tests + format check

## Development

- Rust crate: `rust/`
- TypeScript entrypoints: `src/index.ts`, `src/cli.ts`, `src/native.ts`
- Tests: `tests/index.test.ts`, `tests/cli.test.ts`, `tests/real-conversion.test.ts`

## Release

GitHub Actions runs package checks on every branch push and pull request. The
publish workflow runs only when a version tag matching `v*.*.*` is pushed, for
example `v0.1.0`.

The npm package name is `@alexsun-top/office2pdf`. For the first manual public
publish of this scoped package, run the same local checks and publish with
explicit public access:

```bash
npm ci
npm run check
npm run build
npm publish --access public
```

Do not create or push a release tag until the target npm version, exact tag name,
remote push target, and publish path have been explicitly approved.

The publish workflow is designed for npm Trusted Publishing with GitHub Actions
OIDC. To enable automatic npm authentication without long-lived tokens:

1. Open the package settings on npmjs.com.
2. Add a trusted publisher for `agentsyaml/office2pdf-ts`.
3. Set the trusted workflow file to `.github/workflows/publish.yml`.
4. Allow the workflow to run `npm publish --access public`.

The workflow grants `id-token: write`, so npm can exchange the GitHub OIDC token
during `npm publish --access public`. No `NPM_TOKEN` secret is required for this
path, and no generated `dist/` or `wasm/pkg/` files need to be committed.

If Trusted Publishing is unavailable, create an npm automation token, store it
as the GitHub Actions secret `NPM_TOKEN`, configure `actions/setup-node` with
`registry-url: https://registry.npmjs.org`, and pass the secret only to the
publish step as `NODE_AUTH_TOKEN`. Never commit token values.
