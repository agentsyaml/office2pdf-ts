type ZipEntry = {
  path: string;
  data: Uint8Array;
};

function encodeUtf8(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function crc32(data: Uint8Array): number {
  const table: number[] = [];
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

function writeLe(
  view: DataView,
  offset: number,
  value: number,
  bytes: number,
): void {
  for (let i = 0; i < bytes; i += 1) {
    view.setUint8(offset + i, (value >>> (8 * i)) & 0xff);
  }
}

function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const centralDirectory: Uint8Array[] = [];
  const chunks: Uint8Array[] = [];
  let offset = 0;
  let totalEntries = 0;

  const dosTime = 0;
  const dosDate = 0;

  for (const { path, data } of entries) {
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

    const localHeaderOffset = offset;
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
    writeLe(centralView, 42, localHeaderOffset, 4);
    centralDirectory.push(new Uint8Array(centralHeader));
    centralDirectory.push(nameBytes);
    totalEntries += 1;
  }

  const centralSize = centralDirectory.reduce(
    (acc, entry) => acc + entry.byteLength,
    0,
  );
  const centralOffset = offset;

  const endHeader = new ArrayBuffer(22);
  const endView = new DataView(endHeader);
  endView.setUint32(0, 0x06054b50, true);
  writeLe(endView, 4, 0, 2);
  writeLe(endView, 6, 0, 2);
  writeLe(endView, 8, totalEntries, 2);
  writeLe(endView, 10, totalEntries, 2);
  writeLe(endView, 12, centralSize, 4);
  writeLe(endView, 16, centralOffset, 4);
  writeLe(endView, 20, 0, 2);

  const fullLength =
    chunks.reduce((acc, chunk) => acc + chunk.byteLength, 0) +
    centralSize +
    endHeader.byteLength;
  const result = new Uint8Array(fullLength);
  let cursor = 0;
  for (const chunk of [
    ...chunks,
    ...centralDirectory,
    new Uint8Array(endHeader),
  ]) {
    result.set(chunk, cursor);
    cursor += chunk.byteLength;
  }

  return result;
}

function buildDocxEntries(): ZipEntry[] {
  return [
    {
      path: "[Content_Types].xml",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n</Types>`,
      ),
    },
    {
      path: "_rels/.rels",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "word/document.xml",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/2006/word\" xmlns:ns0=\"http://schemas.openxmlformats.org/drawingml/2006/main\" mc:Ignorable=\"w14 w15 w16se w16cid w16\"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>`,
      ),
    },
    {
      path: "word/styles.xml",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>`,
      ),
    },
    {
      path: "word/_rels/document.xml.rels",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n</Relationships>`,
      ),
    },
  ];
}

function buildPptxEntries(): ZipEntry[] {
  return [
    {
      path: "[Content_Types].xml",
      data: encodeUtf8(
        `<?xml version="1.0" encoding="UTF-8"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>\n  <Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n  <Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>\n  <Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>\n  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n</Types>`,
      ),
    },
    {
      path: "_rels/.rels",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "ppt/presentation.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/></p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"5143500\" type=\"screen16x9\"/></p:presentation>`,
      ),
    },
    {
      path: "ppt/_rels/presentation.xml.rels",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "ppt/slides/slide1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id=\"1\" name=\"Title 1\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
      ),
    },
    {
      path: "ppt/slides/_rels/slide1.xml.rels",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "ppt/slideMasters/slideMaster1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:spTree/></p:cSld></p:sldMaster>`,
      ),
    },
    {
      path: "ppt/slideLayouts/slideLayout1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" type=\"title\"><p:cSld><p:spTree/></p:cSld><p:clrMapOvr/></p:sldLayout>`,
      ),
    },
    {
      path: "ppt/theme/theme1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"></a:theme>`,
      ),
    },
  ];
}

function buildXlsxEntries(): ZipEntry[] {
  return [
    {
      path: "[Content_Types].xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n  <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n</Types>`,
      ),
    },
    {
      path: "_rels/.rels",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "docProps/core.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n  <dc:creator>office2pdf-ts</dc:creator>\n  <dcterms:created xsi:type=\"dcterms:W3CDTF\">2026-06-13T00:00:00Z</dcterms:created>\n  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">2026-06-13T00:00:00Z</dcterms:modified>\n</cp:coreProperties>`,
      ),
    },
    {
      path: "docProps/app.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n  <Application>office2pdf-ts</Application>\n  <DocSecurity>0</DocSecurity>\n  <ScaleCrop>false</ScaleCrop>\n  <Company></Company>\n  <LinksUpToDate>false</LinksUpToDate>\n  <SharedDoc>false</SharedDoc>\n  <HyperlinksChanged>false</HyperlinksChanged>\n  <AppVersion>1.0</AppVersion>\n</Properties>`,
      ),
    },
    {
      path: "xl/workbook.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  <workbookPr defaultThemeVersion=\"160000\"/>\n  <bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"12000\" windowHeight=\"9000\"/></bookViews>\n  <sheets>\n    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n  </sheets>\n</workbook>`,
      ),
    },
    {
      path: "xl/_rels/workbook.xml.rels",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\n  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\n</Relationships>`,
      ),
    },
    {
      path: "xl/worksheets/sheet1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  <dimension ref=\"A1\"/>\n  <sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>\n  <sheetData>\n    <row r=\"1\"><c r=\"A1\" s=\"0\" t=\"s\"><v>0</v></c></row>\n  </sheetData>\n</worksheet>`,
      ),
    },
    {
      path: "xl/styles.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n  <fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/></font></fonts>\n  <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\n  <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>\n</styleSheet>`,
      ),
    },
    {
      path: "xl/sharedStrings.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><t>Hello</t></si></sst>`,
      ),
    },
    {
      path: "xl/theme/theme1.xml",
      data: encodeUtf8(
        `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">\n  <a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:accent1><a:srgbClr val=\"1F497D\"/></a:accent1><a:accent2><a:srgbClr val=\"4F81BD\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
      ),
    },
  ];
}

export type OfficeFormat = "docx" | "pptx" | "xlsx";

export function createOfficeZipBytes(format: OfficeFormat): Uint8Array {
  const entries =
    format === "docx"
      ? buildDocxEntries()
      : format === "pptx"
        ? buildPptxEntries()
        : buildXlsxEntries();

  return writeZip(entries);
}

export function createDocxBytes(): Uint8Array {
  return createOfficeZipBytes("docx");
}

export function createPptxBytes(): Uint8Array {
  return createOfficeZipBytes("pptx");
}

export function createXlsxBytes(): Uint8Array {
  return createOfficeZipBytes("xlsx");
}
