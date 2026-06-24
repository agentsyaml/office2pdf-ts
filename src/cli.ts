#!/usr/bin/env node

import { convertPath, __version__, type ConvertOptions } from "./index.js";
import { mkdir, writeFile } from "node:fs/promises";
import { realpathSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { argv } from "node:process";
import { basename, dirname, extname, join } from "node:path";

const HELP_TEXT = [
  "office2pdf",
  "Usage:",
  "  office2pdf <input.docx|input.pptx|input.xlsx> <output.pdf>",
  "  office2pdf <input.docx|input.pptx|input.xlsx> -o <output.pdf>",
  "  office2pdf <input.docx|input.pptx|input.xlsx> --outdir <directory>",
  "Options:",
  "  -o, --output <path>       Write PDF to the given path",
  "  --outdir <directory>      Write <input basename>.pdf inside directory",
  "  --paper <size>            Paper size such as A4, Letter, or Legal",
  "  --landscape               Render landscape pages",
  "  --pdf-a                   Request PDF/A-2b output",
  "  --sheets <names>          Comma-separated XLSX sheet names",
  "  --slides <range>          PPTX slide range such as 1-3",
  "  --font-path <path>        Font path, repeatable or comma-separated",
  "  office2pdf --help",
  "  office2pdf --version",
].join("\n");

interface ParsedArgs {
  readonly inputPath: string;
  readonly outputPath: string;
  readonly options: ConvertOptions;
}

function printError(message: string): void {
  process.stderr.write(`${message}\n`);
}

function readFlagValue(
  args: readonly string[],
  index: number,
  flag: string,
): string {
  const value = args[index + 1];
  if (value === undefined) {
    throw new Error(`${flag} requires a value`);
  }

  return value;
}

function parseList(value: string, flag: string): string[] {
  const entries = value
    .split(",")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);

  if (entries.length === 0) {
    throw new Error(`${flag} requires at least one value`);
  }

  return entries;
}

function outputPathFromOutdir(inputPath: string, outdir: string): string {
  const inputName = basename(inputPath, extname(inputPath));
  return join(outdir, `${inputName}.pdf`);
}

function resolveOutputPath(
  inputPath: string,
  positionalOutput: string | null,
  flagOutput: string | null,
  outdir: string | null,
): string {
  const outputCount = [positionalOutput, flagOutput, outdir].filter(
    (value) => value !== null,
  ).length;

  if (outputCount === 0) {
    throw new Error("expected input path and output path");
  }

  if (outputCount > 1) {
    throw new Error(
      "output path is ambiguous; provide either positional output, --output, or --outdir",
    );
  }

  if (positionalOutput !== null) {
    return positionalOutput;
  }

  if (flagOutput !== null) {
    return flagOutput;
  }

  if (outdir !== null) {
    return outputPathFromOutdir(inputPath, outdir);
  }

  throw new Error("expected input path and output path");
}

function parseArgs(args: readonly string[]): ParsedArgs {
  const positional: string[] = [];
  const fontPaths: string[] = [];
  const options: ConvertOptions = {};
  let flagOutput: string | null = null;
  let outdir: string | null = null;

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === undefined) {
      continue;
    }

    switch (arg) {
      case "-o":
      case "--output":
        flagOutput = readFlagValue(args, index, arg);
        index += 1;
        break;
      case "--outdir":
        outdir = readFlagValue(args, index, arg);
        index += 1;
        break;
      case "--paper":
        options.paperSize = readFlagValue(args, index, arg);
        index += 1;
        break;
      case "--landscape":
        options.landscape = true;
        break;
      case "--pdf-a":
        options.pdfStandard = "pdf/a-2b";
        break;
      case "--sheets":
        options.sheetNames = parseList(readFlagValue(args, index, arg), arg);
        index += 1;
        break;
      case "--slides":
        options.slideRange = readFlagValue(args, index, arg);
        index += 1;
        break;
      case "--font-path":
        fontPaths.push(...parseList(readFlagValue(args, index, arg), arg));
        index += 1;
        break;
      default:
        if (arg.startsWith("-")) {
          throw new Error(`unknown option: ${arg}`);
        }
        positional.push(arg);
        break;
    }
  }

  if (fontPaths.length > 0) {
    options.fontPaths = fontPaths;
  }

  const [inputPath, positionalOutput, extra] = positional;
  if (!inputPath || extra !== undefined) {
    throw new Error("expected input path and output path");
  }

  return {
    inputPath,
    outputPath: resolveOutputPath(
      inputPath,
      positionalOutput ?? null,
      flagOutput,
      outdir,
    ),
    options,
  };
}

export async function run(
  args: readonly string[] = argv.slice(2),
): Promise<void> {
  if (args.length === 0) {
    printError(HELP_TEXT);
    process.exitCode = 1;
    return;
  }

  const [first] = args;

  if (first === "--help" || first === "-h" || first === "help") {
    process.stdout.write(`${HELP_TEXT}\n`);
    return;
  }

  if (first === "--version" || first === "-v" || first === "version") {
    process.stdout.write(`${__version__}\n`);
    return;
  }

  try {
    const parsed = parseArgs(args);
    const result = await convertPath(parsed.inputPath, parsed.options);
    await mkdir(dirname(parsed.outputPath), { recursive: true });
    await writeFile(parsed.outputPath, result.pdf);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "conversion failed with an unknown error";
    printError(message);
    process.exitCode = 1;
  }
}

function isDirectCliInvocation(): boolean {
  const scriptPath = fileURLToPath(import.meta.url);
  const argvPath = argv[1];

  if (!argvPath) {
    return false;
  }

  if (scriptPath === argvPath) {
    return true;
  }

  try {
    return realpathSync(scriptPath) === realpathSync(argvPath);
  } catch (error) {
    if (error instanceof Error) {
      return false;
    }

    throw error;
  }
}

if (isDirectCliInvocation()) {
  void run();
}
