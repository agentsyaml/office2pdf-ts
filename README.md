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
  convertPath,
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
```

### API

- `__version__: string`
- `Format: { DOCX: "docx"; PPTX: "pptx"; XLSX: "xlsx"; docx: "docx"; pptx: "pptx"; xlsx: "xlsx" }`
- `inferFormat(pathOrUrl: string | URL): "docx" | "pptx" | "xlsx"`
- `convertBytes(data, format, options?)`
- `convertPath(pathOrUrl, options?)`

`ConvertOptions` accepts both camelCase and snake_case aliases:

- `sheetFilter` / `sheet_filter`
- `slideRange` / `slide_range`
- `paperSize` / `paper_size`
- `fontPaths` / `font_paths`
- `pdfStandard` / `pdf_standard`
- `includeWarnings` / `include_warnings`

Unsupported by upstream `office2pdf 0.6`:

- `pageRange` / `page_range` -> throws
  - `page_range is not supported by upstream office2pdf 0.6`
- `memoryLimitMb` / `memory_limit_mb` -> throws
  - `memory_limit_mb is not supported by upstream office2pdf 0.6`

#### Result shape

```ts
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
```

## Scripts

- `npm run wasm:build` — build wasm artifacts to `wasm/pkg`
- `npm run build` — build wasm (release), copy publishable wasm files, and compile TypeScript
- `npm run typecheck` — TypeScript typecheck
- `npm test` — run Vitest
- `npm run check` — typecheck + tests + format check

## Development

- Rust crate: `rust/`
- TypeScript entrypoints: `src/index.ts`, `src/cli.ts`, `src/native.ts`
- Tests: `tests/index.test.ts`

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
