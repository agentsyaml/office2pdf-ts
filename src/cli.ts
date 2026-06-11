#!/usr/bin/env node

import { convertPath, __version__ } from "./index.js";
import { writeFile } from "node:fs/promises";
import { argv, cwd } from "node:process";

const HELP_TEXT = [
  "office2pdf",
  "Usage:",
  "  office2pdf <input.docx|input.pptx|input.xlsx> <output.pdf>",
  "  office2pdf --help",
  "  office2pdf --version",
].join("\n");

function printError(message: string): void {
  process.stderr.write(`${message}\n`);
}

export async function run(
  args: readonly string[] = argv.slice(2),
): Promise<void> {
  if (args.length === 0) {
    printError(HELP_TEXT);
    process.exitCode = 1;
    return;
  }

  const [first, second] = args;

  if (first === "--help" || first === "-h" || first === "help") {
    process.stdout.write(`${HELP_TEXT}\n`);
    return;
  }

  if (first === "--version" || first === "-v" || first === "version") {
    process.stdout.write(`${__version__}\n`);
    return;
  }

  if (args.length !== 2) {
    printError("expected input path and output path");
    printError(HELP_TEXT);
    process.exitCode = 1;
    return;
  }

  if (!first || !second) {
    printError("expected input path and output path");
    printError(HELP_TEXT);
    process.exitCode = 1;
    return;
  }

  try {
    const result = await convertPath(first);
    await writeFile(second, result.pdf);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "conversion failed with an unknown error";
    printError(message);
    process.exitCode = 1;
  }
}

if (import.meta.url === new URL(argv[1] ?? "", `file://${cwd()}/`).href) {
  void run();
}
