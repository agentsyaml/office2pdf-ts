import { execSync } from "node:child_process";
import { mkdtemp, rm, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const repositoryRoot = process.cwd();

function encodeUtf8(value) {
  return new TextEncoder().encode(value);
}

function crc32(data) {
  const table = [];
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c;
  }

  let checksum = 0xffffffff;
  for (let i = 0; i < data.length; i += 1) {
    checksum ^= data[i] ?? 0;
    for (let k = 0; k < 8; k += 1) {
      checksum = checksum & 1 ? 0xedb88320 ^ (checksum >>> 1) : checksum >>> 1;
    }
  }

  return (checksum ^ 0xffffffff) >>> 0;
}

function writeLe(dataView, offset, value, bytes) {
  for (let i = 0; i < bytes; i += 1) {
    dataView.setUint8(offset + i, (value >>> (8 * i)) & 0xff);
  }
}

function writeZip(entries) {
  const chunks = [];
  const centralDirectory = [];
  let offset = 0;
  let totalEntries = 0;

  const dosTime = 0;
  const dosDate = 0;

  for (const [path, data] of entries) {
    const nameBytes = encodeUtf8(path);
    const crc = crc32(data);

    const localHeader = new ArrayBuffer(30);
    const localView = new DataView(localHeader);
    localView.setUint32(0, 0x04034b50, true);
    writeLe(localView, 4, 20, 2);
    writeLe(localView, 6, 0, 2);
    writeLe(localView, 8, 0, 2);
    writeLe(localView, 10, dosTime, 2);
    writeLe(localView, 12, dosDate, 2);
    writeLe(localView, 14, crc, 4);
    writeLe(localView, 18, data.length, 4);
    writeLe(localView, 22, data.length, 4);
    writeLe(localView, 26, nameBytes.length, 2);
    writeLe(localView, 28, 0, 2);
    chunks.push(new Uint8Array(localHeader));
    chunks.push(nameBytes);
    chunks.push(data);

    const localOffset = offset;
    offset += localHeader.byteLength + nameBytes.length + data.length;

    const centralHeader = new ArrayBuffer(46);
    const centralView = new DataView(centralHeader);
    centralView.setUint32(0, 0x02014b50, true);
    writeLe(centralView, 4, 20, 2);
    writeLe(centralView, 6, 20, 2);
    writeLe(centralView, 8, 0, 2);
    writeLe(centralView, 10, 0, 2);
    writeLe(centralView, 12, dosTime, 2);
    writeLe(centralView, 14, dosDate, 2);
    writeLe(centralView, 16, crc, 4);
    writeLe(centralView, 20, data.length, 4);
    writeLe(centralView, 24, data.length, 4);
    writeLe(centralView, 28, nameBytes.length, 2);
    writeLe(centralView, 30, 0, 2);
    writeLe(centralView, 32, 0, 2);
    writeLe(centralView, 34, 0, 2);
    writeLe(centralView, 36, 0, 2);
    writeLe(centralView, 38, 0, 4);
    writeLe(centralView, 42, localOffset, 4);

    centralDirectory.push(new Uint8Array(centralHeader));
    centralDirectory.push(nameBytes);
    totalEntries += 1;
  }

  const centralSize = centralDirectory.reduce(
    (sum, item) => sum + item.byteLength,
    0,
  );
  const centralOffset = offset;

  const endOfCentral = new ArrayBuffer(22);
  const endView = new DataView(endOfCentral);
  endView.setUint32(0, 0x06054b50, true);
  writeLe(endView, 4, 0, 2);
  writeLe(endView, 6, 0, 2);
  writeLe(endView, 8, totalEntries, 2);
  writeLe(endView, 10, totalEntries, 2);
  writeLe(endView, 12, centralSize, 4);
  writeLe(endView, 16, centralOffset, 4);
  writeLe(endView, 20, 0, 2);

  const output = new Uint8Array(
    chunks.reduce((sum, item) => sum + item.byteLength, 0) +
      centralSize +
      endOfCentral.byteLength,
  );
  let cursor = 0;
  for (const chunk of [
    ...chunks,
    ...centralDirectory,
    new Uint8Array(endOfCentral),
  ]) {
    output.set(chunk, cursor);
    cursor += chunk.byteLength;
  }

  return output;
}

function createDocxBytes() {
  const entries = [
    [
      "[Content_Types].xml",
      encodeUtf8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    ],
    [
      "_rels/.rels",
      encodeUtf8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    ],
    [
      "word/document.xml",
      encodeUtf8(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/2006/word" xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" mc:Ignorable="w14 w15 w16se w16cid w16"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>`),
    ],
    [
      "word/styles.xml",
      encodeUtf8(`<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    ],
    [
      "word/_rels/document.xml.rels",
      encodeUtf8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    ],
  ];

  return writeZip(entries);
}

function assertPdf(filePath) {
  return readFile(filePath).then((data) => {
    const header = new TextDecoder().decode(data.subarray(0, 5));
    if (header !== "%PDF-") {
      throw new Error(`invalid PDF header in ${filePath}`);
    }
  });
}

function runCommand(command, options = {}) {
  return execSync(command, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "inherit"],
    shell: true,
    ...options,
  }).trim();
}

async function main() {
  const packageJson = JSON.parse(
    await readFile(join(repositoryRoot, "package.json"), "utf8"),
  );
  const expectedVersion = packageJson.version;

  const packResult = JSON.parse(
    runCommand("npm pack --json", {
      cwd: repositoryRoot,
    }),
  );
  const packageTarball = join(repositoryRoot, packResult[0].filename);

  const workspace = await mkdtemp(
    join(tmpdir(), "office2pdf-installed-smoke-"),
  );
  try {
    runCommand("npm init -y", { cwd: workspace });
    runCommand(`npm install --silent ${packageTarball}`, { cwd: workspace });

    const apiImportUrl = new URL(
      "./node_modules/@alexsun-top/office2pdf/dist/index.js",
      `file://${workspace}/`,
    );
    const imported = await import(apiImportUrl);

    const cliCommand = join(workspace, "node_modules", ".bin", "office2pdf");
    const versionOutput = runCommand(`"${cliCommand}" --version`).trim();
    if (versionOutput !== expectedVersion) {
      throw new Error(`unexpected CLI version ${versionOutput}`);
    }

    const inputPath = join(workspace, "sample.docx");
    const outputPath = join(workspace, "sample.pdf");
    await writeFile(inputPath, Buffer.from(createDocxBytes()));

    runCommand(`"${cliCommand}" "${inputPath}" "${outputPath}"`);
    await assertPdf(outputPath);

    const apiResult = await imported.convertBytes(
      await readFile(inputPath),
      imported.Format.docx,
    );
    if (!(apiResult?.pdf instanceof Uint8Array)) {
      throw new Error("API import path did not return Uint8Array payload");
    }

    const apiHeader = new TextDecoder().decode(apiResult.pdf.subarray(0, 5));
    if (apiHeader !== "%PDF-") {
      throw new Error("API import path output is not a PDF");
    }

    console.log("installed-package smoke passed");
  } finally {
    await rm(workspace, { recursive: true, force: true });
    await rm(packageTarball, { force: true });
  }
}

await main();
