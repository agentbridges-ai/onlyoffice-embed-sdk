import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.join(__dirname, "..", "files");
const encoder = new TextEncoder();

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function writeU16(data, offset, value) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
}

function writeU32(data, offset, value) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
  data[offset + 2] = (value >>> 16) & 0xff;
  data[offset + 3] = (value >>> 24) & 0xff;
}

function concatBytes(parts) {
  const output = new Uint8Array(
    parts.reduce((total, part) => total + part.length, 0),
  );
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function xml(strings, ...values) {
  return strings.reduce((result, value, index) => {
    return `${result}${value}${values[index] ?? ""}`;
  }, "").trim();
}

function utf8(value) {
  return encoder.encode(value);
}

function zip(entries) {
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const [name, content] of entries) {
    const nameBytes = utf8(name);
    const data =
      typeof content === "string" ? utf8(content) : new Uint8Array(content);
    const crc = crc32(data);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeU32(localHeader, 0, 0x04034b50);
    writeU16(localHeader, 4, 20);
    writeU16(localHeader, 6, 0);
    writeU16(localHeader, 8, 0);
    writeU16(localHeader, 10, 0);
    writeU16(localHeader, 12, 0);
    writeU32(localHeader, 14, crc);
    writeU32(localHeader, 18, data.length);
    writeU32(localHeader, 22, data.length);
    writeU16(localHeader, 26, nameBytes.length);
    writeU16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeU32(centralHeader, 0, 0x02014b50);
    writeU16(centralHeader, 4, 20);
    writeU16(centralHeader, 6, 20);
    writeU16(centralHeader, 8, 0);
    writeU16(centralHeader, 10, 0);
    writeU16(centralHeader, 12, 0);
    writeU16(centralHeader, 14, 0);
    writeU32(centralHeader, 16, crc);
    writeU32(centralHeader, 20, data.length);
    writeU32(centralHeader, 24, data.length);
    writeU16(centralHeader, 28, nameBytes.length);
    writeU16(centralHeader, 30, 0);
    writeU16(centralHeader, 32, 0);
    writeU16(centralHeader, 34, 0);
    writeU16(centralHeader, 36, 0);
    writeU32(centralHeader, 38, 0);
    writeU32(centralHeader, 42, localOffset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    localOffset += localHeader.length + data.length;
  }

  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, entries.length);
  writeU16(end, 10, entries.length);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, localOffset);

  return concatBytes([...localParts, centralDirectory, end]);
}

const coreProps = xml`
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>OnlyOffice E2E Fixture</dc:title>
    <dc:creator>Playwright</dc:creator>
    <cp:lastModifiedBy>Playwright</cp:lastModifiedBy>
  </cp:coreProperties>
`;

const appProps = xml`
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>OnlyOffice Embed SDK E2E</Application>
  </Properties>
`;

function docxDocumentXml({ invalidBookmark = false, largeText = "" } = {}) {
  const bookmark = invalidBookmark
    ? '<w:bookmarkStart w:id="DingTalkBookmark" w:name="_GoBack"/><w:bookmarkEnd w:id="DingTalkBookmark"/>'
    : "";
  return xml`
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          ${bookmark}
          <w:r><w:t>OnlyOffice generated DOCX fixture</w:t></w:r>
        </w:p>
        <w:p><w:r><w:t>${largeText}</w:t></w:r></w:p>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:body>
    </w:document>
  `;
}

function docx({ invalidBookmark = false, largeText = "" } = {}) {
  return zip([
    [
      "[Content_Types].xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
      `,
    ],
    [
      "_rels/.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
      `,
    ],
    ["docProps/core.xml", coreProps],
    ["docProps/app.xml", appProps],
    ["word/document.xml", docxDocumentXml({ invalidBookmark, largeText })],
    [
      "word/_rels/document.xml.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
      `,
    ],
  ]);
}

function xlsx() {
  const longText =
    "OnlyOffice generated XLSX fixture with long text. ".repeat(20).trim();
  return zip([
    [
      "[Content_Types].xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
      `,
    ],
    [
      "_rels/.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
      `,
    ],
    ["docProps/core.xml", coreProps],
    ["docProps/app.xml", appProps],
    [
      "xl/workbook.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
      `,
    ],
    [
      "xl/_rels/workbook.xml.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
      `,
    ],
    [
      "xl/styles.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
          <borders count="1"><border/></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        </styleSheet>
      `,
    ],
    [
      "xl/worksheets/sheet1.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:C3"/>
          <sheetData>
            <row r="1">
              <c r="A1" t="inlineStr"><is><t>Label</t></is></c>
              <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
              <c r="C1" t="inlineStr"><is><t>Notes</t></is></c>
            </row>
            <row r="2">
              <c r="A2" t="inlineStr"><is><t>Total</t></is></c>
              <c r="B2"><v>42</v></c>
              <c r="C2" t="inlineStr"><is><t>${longText}</t></is></c>
            </row>
            <row r="3">
              <c r="A3" t="inlineStr"><is><t>Formula</t></is></c>
              <c r="B3"><f>B2*2</f><v>84</v></c>
            </row>
          </sheetData>
        </worksheet>
      `,
    ],
  ]);
}

const fixtures = [
  {
    name: "edge-invalid-bookmark.docx",
    data: docx({ invalidBookmark: true }),
    kind: "positive",
    fileType: "DOCX",
    source: "#29 / PR #30 invalid DingTalk DOCX bookmark",
  },
  {
    name: "xml-limit.docx",
    data: docx({ largeText: "XML_LIMIT ".repeat(4096) }),
    kind: "negative",
    fileType: "DOCX",
    source: "#34 configurable Office XML size guard",
  },
  {
    name: "mismatch-xlsx-as-docx.docx",
    data: xlsx(),
    kind: "negative",
    fileType: "DOCX",
    source: "extension/content mismatch rejection",
  },
  {
    name: "plain-text-as-docx.docx",
    data: utf8("dsadads\ndasdas"),
    kind: "positive",
    fileType: "DOCX",
    source: "plain text content with DOCX extension fallback",
  },
  {
    name: "plain-text-as-xlsx.xlsx",
    data: utf8("dsadads.dasdas\n"),
    kind: "positive",
    fileType: "XLSX",
    source: "plain text content with XLSX extension fallback",
  },
];

fs.rmSync(outputDir, { recursive: true, force: true });
fs.mkdirSync(outputDir, { recursive: true });

for (const fixture of fixtures) {
  fs.writeFileSync(path.join(outputDir, fixture.name), fixture.data);
}

fs.writeFileSync(
  path.join(outputDir, "manifest.json"),
  JSON.stringify(
    fixtures.map(({ data, ...fixture }) => ({
      ...fixture,
      size: data.length,
    })),
    null,
    2,
  ),
);

console.log(`Generated ${fixtures.length} Office fixtures in ${outputDir}`);
