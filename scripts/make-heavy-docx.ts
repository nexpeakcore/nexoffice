/**
 * Generate a large DOCX for performance measurement.
 *
 * The point is a fixture whose cost is known and repeatable: an explicit page
 * break ends every page, so the page count is exactly what was asked for
 * rather than whatever the layout engine happens to produce, and a seeded PRNG
 * makes the same arguments yield byte-identical output. Images are distinct
 * per occurrence — a repeated image would be resolved once and cached, which
 * is the opposite of what an image-cache measurement needs.
 *
 * Run: bun scripts/make-heavy-docx.ts --pages 2000
 */

import JSZip from 'jszip';
import { deflateSync } from 'node:zlib';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const ZIP_DATE = new Date('2026-01-01T00:00:00Z');

interface Options {
  pages: number;
  out: string;
  imageEvery: number;
  tableEvery: number;
  paragraphsPerPage: number;
  seed: number;
  mixedScripts: boolean;
}

function parseArgs(argv: string[]): Options {
  const flags = new Map<string, string>();
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith('--')) continue;
    const key = arg.slice(2);
    const next = argv[i + 1];
    if (next && !next.startsWith('--')) {
      flags.set(key, next);
      i += 1;
    } else {
      flags.set(key, 'true');
    }
  }
  const number = (key: string, fallback: number): number => {
    const raw = flags.get(key);
    if (raw === undefined) return fallback;
    const value = Number(raw);
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`--${key} must be a non-negative number, got ${raw}`);
    }
    return Math.floor(value);
  };
  const pages = Math.max(1, number('pages', 2000));
  return {
    pages,
    out: flags.get('out') ?? path.join(ROOT, 'test-fixtures', `heavy-${pages}p.docx`),
    // 0 disables. Defaults put ~80 images and ~200 tables in a 2000-page file:
    // enough to overrun the 64MB image cache and to make table layout a real
    // share of the total, without making the document mostly furniture.
    imageEvery: number('image-every', 25),
    tableEvery: number('table-every', 10),
    paragraphsPerPage: Math.max(1, number('paragraphs-per-page', 7)),
    seed: number('seed', 20260827),
    mixedScripts: flags.get('mixed') === 'true',
  };
}

/** mulberry32 — small, seeded, and stable across runtimes. */
function makeRandom(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

const LATIN_WORDS = `the layout engine measures every run before it places a single line because
shaping decides width and width decides where the break falls document parts arrive as
compressed streams that must be inflated parsed and validated in order paragraph properties
inherit from styles which inherit from document defaults so resolving one run can touch four
tables of state a page is finished only when its floating objects footnotes and header bands
have all been placed and none of them can be placed until the text they anchor to has been
measured which is why pagination is a fixed point rather than a single pass`
  .split(/\s+/)
  .filter(Boolean);

// Diacritics and CJK force the fallback chain and the shaper down paths plain
// ASCII never reaches, which is where a mixed-script document gets expensive.
const VIETNAMESE_WORDS =
  'trình bày tài liệu được đo đạc từng đoạn trước khi xuống dòng vì chiều rộng quyết định điểm ngắt bảng biểu và chú thích cuối trang đều phải xếp xong mới coi là hoàn tất một trang'.split(
    /\s+/
  );
const CJK_WORDS = '排版 引擎 在 放置 每 一行 之前 都会 测量 文本 宽度 决定 换行 位置'.split(/\s+/);

function sentence(random: () => number, pool: string[], words: number): string {
  const picked: string[] = [];
  for (let i = 0; i < words; i += 1) {
    picked.push(pool[Math.floor(random() * pool.length)] ?? 'text');
  }
  const text = picked.join(' ');
  return `${text.charAt(0).toUpperCase()}${text.slice(1)}.`;
}

function bodyText(random: () => number, mixed: boolean): string {
  const sentences: string[] = [];
  const count = 3 + Math.floor(random() * 3);
  for (let i = 0; i < count; i += 1) {
    const pool = !mixed
      ? LATIN_WORDS
      : random() < 0.15
        ? VIETNAMESE_WORDS
        : random() < 0.08
          ? CJK_WORDS
          : LATIN_WORDS;
    sentences.push(sentence(random, pool, 10 + Math.floor(random() * 14)));
  }
  return sentences.join(' ');
}

// ---------------------------------------------------------------------------
// Images — a minimal PNG encoder, so each occurrence is genuinely its own bytes
// ---------------------------------------------------------------------------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const byte of bytes) c = CRC_TABLE[(c ^ byte) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i += 1) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/** An RGBA PNG whose pixels depend on `index`, so no two are alike. */
function makePng(index: number, width: number, height: number): Uint8Array {
  const raw = new Uint8Array(height * (1 + width * 4));
  const hue = (index * 47) % 360;
  for (let y = 0; y < height; y += 1) {
    const row = y * (1 + width * 4);
    raw[row] = 0;
    for (let x = 0; x < width; x += 1) {
      const p = row + 1 + x * 4;
      raw[p] = (hue + x * 0.3) % 256;
      raw[p + 1] = (y * 0.4 + index * 13) % 256;
      raw[p + 2] = (x * 0.2 + y * 0.2 + index * 29) % 256;
      raw[p + 3] = 255;
    }
  }
  const ihdr = new Uint8Array(13);
  const view = new DataView(ihdr.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const idat = new Uint8Array(deflateSync(raw, { level: 6 }));
  const signature = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const parts = [
    signature,
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const png = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    png.set(part, offset);
    offset += part.length;
  }
  return png;
}

// ---------------------------------------------------------------------------
// Body parts
// ---------------------------------------------------------------------------

const IMAGE_WIDTH = 900;
const IMAGE_HEIGHT = 600;
/** 914400 EMU per inch; 4.5in wide keeps the image inside a 6.5in text column. */
const IMAGE_CX = Math.round(4.5 * 914400);
const IMAGE_CY = Math.round((4.5 * IMAGE_HEIGHT / IMAGE_WIDTH) * 914400);

function para(style: string, text: string, extra = ''): string {
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/>${extra}</w:pPr><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

function pageBreak(): string {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

function table(random: () => number, page: number): string {
  const columns = 5;
  const rows = 8;
  const width = Math.floor(9360 / columns);
  const grid = Array.from({ length: columns }, () => `<w:gridCol w:w="${width}"/>`).join('');
  const cell = (text: string, header: boolean): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${header ? '<w:shd w:val="clear" w:color="auto" w:fill="DCE6F1"/>' : ''}</w:tcPr>` +
    `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r>${header ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p></w:tc>`;
  const body: string[] = [];
  for (let r = 0; r < rows; r += 1) {
    const cells: string[] = [];
    for (let c = 0; c < columns; c += 1) {
      cells.push(
        r === 0
          ? cell(`Column ${c + 1}`, true)
          : cell(
              c === 0
                ? `Row ${r} · page ${page}`
                : (random() * 10000).toFixed(c === 1 ? 0 : 2),
              false
            )
      );
    }
    body.push(`<w:tr>${cells.join('')}</w:tr>`);
  }
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>`)
      .join('') +
    '</w:tblBorders>' +
    '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>' +
    '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>' +
    '</w:tblPr>' +
    `<w:tblGrid>${grid}</w:tblGrid>${body.join('')}</w:tbl>`
  );
}

function image(relId: string, index: number): string {
  return (
    '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${IMAGE_CX}" cy="${IMAGE_CY}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${index + 1}" name="Figure ${index + 1}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="figure-${index + 1}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IMAGE_CX}" cy="${IMAGE_CY}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  );
}

// ---------------------------------------------------------------------------
// Fixed parts
// ---------------------------------------------------------------------------

function contentTypes(hasImages: boolean): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasImages ? '<Default Extension="png" ContentType="image/png"/>' : ''}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;
}

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

function coreXml(pages: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>NexOffice performance fixture — ${pages} pages</dc:title>
  <dc:creator>make-heavy-docx</dc:creator>
  <cp:lastModifiedBy>make-heavy-docx</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="2E74B5"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="200"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
  </w:style>
</w:styles>`;

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr><w:t>NexOffice performance fixture</w:t></w:r></w:p>
</w:hdr>`;

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
    <w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>`;

// ---------------------------------------------------------------------------
// Assemble
// ---------------------------------------------------------------------------

const options = parseArgs(process.argv.slice(2));
const random = makeRandom(options.seed);

const media: Array<{ name: string; bytes: Uint8Array; relId: string }> = [];
const body: string[] = [];
let paragraphs = 0;
let tables = 0;

for (let page = 1; page <= options.pages; page += 1) {
  body.push(para(page % 50 === 1 ? 'Heading1' : 'Heading2', `Section ${page}`));
  paragraphs += 1;

  if (options.imageEvery > 0 && page % options.imageEvery === 0) {
    const index = media.length;
    const relId = `rIdImg${index}`;
    media.push({
      name: `figure-${index + 1}.png`,
      bytes: makePng(index, IMAGE_WIDTH, IMAGE_HEIGHT),
      relId,
    });
    body.push(image(relId, index));
    body.push(para('Caption', `Figure ${index + 1}. Generated plate on page ${page}.`));
    paragraphs += 2;
  }

  if (options.tableEvery > 0 && page % options.tableEvery === 0) {
    body.push(table(random, page));
    tables += 1;
  }

  for (let i = 0; i < options.paragraphsPerPage; i += 1) {
    body.push(para('Normal', bodyText(random, options.mixedScripts)));
    paragraphs += 1;
  }

  // The break belongs to the page it ends, so the last page does not open an
  // empty one after it.
  if (page < options.pages) body.push(pageBreak());
}

const documentRels = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
  ...media.map(
    (entry) =>
      `<Relationship Id="${entry.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${entry.name}"/>`
  ),
  '</Relationships>',
].join('');

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>${body.join('')}<w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
      <w:footerReference w:type="default" r:id="rId3"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr></w:body>
</w:document>`;

const zip = new JSZip();
const opts = { date: ZIP_DATE, createFolders: false };
zip.file('[Content_Types].xml', contentTypes(media.length > 0), opts);
zip.file('_rels/.rels', RELS_XML, opts);
zip.file('docProps/core.xml', coreXml(options.pages), opts);
zip.file('word/_rels/document.xml.rels', documentRels, opts);
zip.file('word/document.xml', documentXml, opts);
zip.file('word/styles.xml', STYLES_XML, opts);
zip.file('word/header1.xml', HEADER_XML, opts);
zip.file('word/footer1.xml', FOOTER_XML, opts);
for (const entry of media) zip.file(`word/media/${entry.name}`, entry.bytes, opts);

const buffer = await zip.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 6 },
});

fs.mkdirSync(path.dirname(options.out), { recursive: true });
fs.writeFileSync(options.out, buffer);

const mb = (bytes: number): string => `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
const mediaBytes = media.reduce((sum, entry) => sum + entry.bytes.length, 0);
console.log(`Wrote ${options.out}`);
console.log(`  pages          ${options.pages}`);
console.log(`  paragraphs     ${paragraphs}`);
console.log(`  tables         ${tables}`);
console.log(`  images         ${media.length} (${mb(mediaBytes)} encoded, ${mb(media.length * IMAGE_WIDTH * IMAGE_HEIGHT * 4)} decoded)`);
console.log(`  document.xml   ${mb(Buffer.byteLength(documentXml))} uncompressed`);
console.log(`  file           ${mb(buffer.length)}`);
