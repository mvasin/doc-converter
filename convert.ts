#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { spawn } from "node:child_process";
import mammoth from "mammoth";
import TurndownService from "turndown";
import { PDFParse } from "pdf-parse";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";

const supportedExtensions = new Set([
  ".docx",
  ".DOCX",
  ".doc",
  ".DOC",
  ".pdf",
  ".PDF",
]);

const turndownService = new TurndownService({ headingStyle: "atx" });
turndownService.addRule("removeImages", {
  filter: "img",
  replacement: () => "",
});

const argv = yargs(hideBin(process.argv))
  .option("input", {
    type: "string",
    demandOption: true,
    describe: "Path to the input file or directory",
  })
  .option("output", {
    type: "string",
    demandOption: true,
    describe: "Path to the output file or directory",
  })
  .option("clear-output", {
    type: "boolean",
    default: false,
    describe:
      "Remove existing files in the output directory before generating Markdown",
  })
  .strict()
  .parseSync();

const inputPath = path.resolve(process.cwd(), argv.input!);
const outputPath = path.resolve(process.cwd(), argv.output!);
const clearOutput = Boolean(argv["clear-output"]);

async function convertDocxToMarkdown(filePath: string): Promise<string> {
  const { value: html } = await mammoth.convertToHtml({ path: filePath });
  return turndownService.turndown(html);
}

async function runLibreOfficeConversion(
  filePath: string,
  outDir: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const args = [
      "--headless",
      "--convert-to",
      "docx",
      "--outdir",
      outDir,
      filePath,
    ];
    const proc = spawn("soffice", args, { stdio: "ignore" });

    proc.on("error", (error) => reject(error));
    proc.on("exit", (code, signal) => {
      if (code === 0) {
        resolve();
      } else {
        const reason = signal ? `signal ${signal}` : `exit code ${code}`;
        reject(new Error(`LibreOffice conversion failed (${reason})`));
      }
    });
  });
}

async function convertDocToDocx(
  filePath: string
): Promise<{ docxPath: string; cleanup: () => Promise<void> }> {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "input-docx-"));
  await runLibreOfficeConversion(filePath, tempDir);

  const docxPath = path.join(
    tempDir,
    `${path.basename(filePath, path.extname(filePath))}.docx`
  );
  await fs.access(docxPath);

  return {
    docxPath,
    cleanup: async () => {
      await fs.rm(tempDir, { recursive: true, force: true });
    },
  };
}

async function convertPdfToMarkdown(filePath: string): Promise<string> {
  const fileBuffer = await fs.readFile(filePath);
  const pdfParse = new PDFParse({ data: fileBuffer });
  const textResult = await pdfParse.getText();
  await pdfParse.destroy();

  // Extract text and format as markdown with basic structure preservation
  const lines = textResult.text
    .split("\n")
    .filter((line: string) => line.trim());

  // Group consecutive lines and clean them up
  return lines
    .map((line: string) => line.trim())
    .filter((line: string) => line.length > 0)
    .join("\n\n"); // Double newlines for readability
}

async function convertDocument(filePath: string): Promise<string> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".docx") {
    return convertDocxToMarkdown(filePath);
  }

  if (ext === ".doc") {
    const { docxPath, cleanup } = await convertDocToDocx(filePath);
    try {
      return await convertDocxToMarkdown(docxPath);
    } finally {
      await cleanup();
    }
  }

  if (ext === ".pdf") {
    return convertPdfToMarkdown(filePath);
  }

  throw new Error(`Unsupported file type: ${filePath}`);
}

async function convertSingleFile(
  inputPath: string,
  outputPath: string,
  clearOutput: boolean
): Promise<void> {
  if (!supportedExtensions.has(path.extname(inputPath))) {
    throw new Error(`Unsupported input file: ${inputPath}`);
  }

  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  try {
    const existing = await fs.stat(outputPath);
    if (existing.isDirectory()) {
      throw new Error(
        "Output path points to a directory but a file was expected"
      );
    }

    if (clearOutput) {
      await fs.rm(outputPath, { force: true });
    } else {
      console.log(`Skipping ${outputPath} because it already exists`);
      return;
    }
  } catch (error: any) {
    if (error.code !== "ENOENT") {
      throw error;
    }
  }

  const markdown = await convertDocument(inputPath);
  await fs.writeFile(outputPath, markdown, "utf8");
  console.log(`✓ Converted ${inputPath} → ${outputPath}`);
}

async function convertDirectory(
  inputDir: string,
  outputDir: string,
  clearOutput: boolean
): Promise<void> {
  if (clearOutput) {
    await fs.rm(outputDir, { recursive: true, force: true });
  }

  await fs.mkdir(outputDir, { recursive: true });
  const entries = await fs.readdir(inputDir, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile())
    .filter((entry) => supportedExtensions.has(path.extname(entry.name)));

  if (files.length === 0) {
    console.warn("No supported files found in", inputDir);
    return;
  }

  for (const entry of files) {
    const sourcePath = path.join(inputDir, entry.name);
    const mdName = `${path.basename(entry.name, path.extname(entry.name))}.md`;
    const targetPath = path.join(outputDir, mdName);

    if (!clearOutput) {
      try {
        await fs.access(targetPath);
        console.log(`Skipping ${targetPath} because it already exists`);
        continue;
      } catch (error: any) {
        if (error.code !== "ENOENT") {
          console.error(`Failed to probe ${targetPath}:`, error);
          continue;
        }
      }
    } else {
      await fs.rm(targetPath, { force: true }).catch(() => undefined);
    }

    try {
      const markdown = await convertDocument(sourcePath);
      await fs.writeFile(targetPath, markdown, "utf8");
      console.log(`✓ Converted ${sourcePath} → ${targetPath}`);
    } catch (error) {
      console.error(`✗ Failed to convert ${sourcePath}:`, error);
    }
  }
}

async function main(): Promise<void> {
  try {
    const stats = await fs.stat(inputPath);

    if (stats.isFile()) {
      await convertSingleFile(inputPath, outputPath, clearOutput);
    } else if (stats.isDirectory()) {
      await convertDirectory(inputPath, outputPath, clearOutput);
    } else {
      throw new Error("--input must point to a file or directory");
    }
  } catch (error) {
    console.error("Failed to process files:", error);
    process.exitCode = 1;
  }
}

await main();
