/**
 * Batch 1 tests — Numbering, Settings, Page Numbering, Section Mechanics,
 *                  Table Cell Completion.
 *
 * Validates:
 * 1. word/numbering.xml is emitted when numbering is defined
 * 2. word/numbering.xml is NOT emitted when no numbering is defined
 * 3. word/settings.xml is emitted when settings are configured
 * 4. word/settings.xml is NOT emitted when no settings are configured
 * 5. w:numPr is serialized in paragraph pPr
 * 6. pgNumType is serialized in sectPr
 * 7. textDirection is serialized in sectPr and paragraph pPr
 * 8. Table cell borders, vMerge, noWrap, tcMar are serialized
 * 9. Mid-document section break via pPr.sectPr
 * 10. Numbering XML structure (abstractNum, num, lvl)
 */

import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder, paragraph, emptyParagraph,
  bulletListLevel, decimalListLevel, arabicAbjadListLevel,
  ptToTwips,
} from '../src/index';

async function buildAndOpen(b: DocxBuilder) {
  const buf = await b.build();
  return JSZip.loadAsync(buf);
}

async function readPart(zip: JSZip, name: string): Promise<string | null> {
  const entry = zip.file(name);
  if (!entry) return null;
  return entry.async('string');
}

function partNames(zip: JSZip): string[] {
  return Object.keys(zip.files).filter(n => !zip.files[n].dir);
}

// ---------------------------------------------------------------------------
// Numbering part emission
// ---------------------------------------------------------------------------

describe('word/numbering.xml emission', () => {
  it('is NOT emitted when no numbering is defined', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['plain']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).not.toContain('word/numbering.xml');
  });

  it('IS emitted when numbering is defined', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [decimalListLevel(0)] });
    b.addBlock(paragraph(['item']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).toContain('word/numbering.xml');
  });

  it('has correct content type in [Content_Types].xml', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [decimalListLevel(0)] });
    b.addBlock(paragraph(['item']));
    const zip = await buildAndOpen(b);
    const ct = await readPart(zip, '[Content_Types].xml');
    expect(ct).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  });

  it('has relationship entry in word/_rels/document.xml.rels', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [decimalListLevel(0)] });
    b.addBlock(paragraph(['item']));
    const zip = await buildAndOpen(b);
    const rels = await readPart(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('relationships/numbering');
    expect(rels).toContain('Target="numbering.xml"');
  });
});

// ---------------------------------------------------------------------------
// Numbering XML structure
// ---------------------------------------------------------------------------

describe('word/numbering.xml structure', () => {
  let xml: string;

  beforeAll(async () => {
    const b = new DocxBuilder();
    b.defineNumbering({
      multiLevelType: 'singleLevel',
      levels: [decimalListLevel(0)],
    });
    const zip = await buildAndOpen(b);
    xml = (await readPart(zip, 'word/numbering.xml'))!;
  });

  it('has WML namespace', () => {
    expect(xml).toContain('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  });

  it('has w:abstractNum element', () => {
    expect(xml).toContain('<w:abstractNum');
    expect(xml).toContain('w:abstractNumId="0"');
  });

  it('has w:multiLevelType', () => {
    expect(xml).toContain('<w:multiLevelType w:val="singleLevel"/>');
  });

  it('has w:lvl with ilvl=0', () => {
    expect(xml).toContain('w:ilvl="0"');
  });

  it('has w:numFmt decimal', () => {
    expect(xml).toContain('<w:numFmt w:val="decimal"/>');
  });

  it('has w:lvlText', () => {
    expect(xml).toContain('<w:lvlText w:val="%1."/>');
  });

  it('has w:num element', () => {
    expect(xml).toContain('<w:num');
    expect(xml).toContain('w:numId="1"');
  });

  it('w:num references abstractNumId=0', () => {
    expect(xml).toContain('<w:abstractNumId w:val="0"/>');
  });
});

// ---------------------------------------------------------------------------
// w:numPr in paragraph
// ---------------------------------------------------------------------------

describe('w:numPr serialization', () => {
  it('emits w:numPr with ilvl and numId', async () => {
    const b = new DocxBuilder();
    const numId = b.defineNumbering({ levels: [decimalListLevel(0)] });
    b.addBlock(paragraph(['item'], { numPr: { numId, ilvl: 0 } }));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:numPr>');
    expect(xml).toContain('<w:ilvl w:val="0"/>');
    expect(xml).toContain(`<w:numId w:val="${numId}"/>`);
  });

  it('multiple numbering definitions get distinct numIds', async () => {
    const b = new DocxBuilder();
    const id1 = b.defineNumbering({ levels: [bulletListLevel(0)] });
    const id2 = b.defineNumbering({ levels: [decimalListLevel(0)] });
    expect(id1).not.toBe(id2);
    expect(id1).toBe(1);
    expect(id2).toBe(2);
  });

  it('bullet list level uses bullet numFmt', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [bulletListLevel(0)] });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/numbering.xml');
    expect(xml).toContain('<w:numFmt w:val="bullet"/>');
  });

  it('arabicAbjad list level uses arabicAbjad numFmt', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [arabicAbjadListLevel(0)] });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/numbering.xml');
    expect(xml).toContain('<w:numFmt w:val="arabicAbjad"/>');
  });
});

// ---------------------------------------------------------------------------
// Settings part
// ---------------------------------------------------------------------------

describe('word/settings.xml emission', () => {
  it('is NOT emitted when no settings configured', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).not.toContain('word/settings.xml');
  });

  it('IS emitted when updateFields is set', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).toContain('word/settings.xml');
  });

  it('has correct content type', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const ct = await readPart(zip, '[Content_Types].xml');
    expect(ct).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  });

  it('has relationship entry', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const rels = await readPart(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('relationships/settings');
    expect(rels).toContain('Target="settings.xml"');
  });

  it('emits w:updateFields', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/settings.xml');
    expect(xml).toContain('<w:updateFields w:val="1"/>');
  });

  it('emits w:evenAndOddHeaders', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ evenAndOddHeaders: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/settings.xml');
    expect(xml).toContain('<w:evenAndOddHeaders/>');
  });

  it('emits w:defaultTabStop', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ defaultTabStop: 720 });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/settings.xml');
    expect(xml).toContain('<w:defaultTabStop w:val="720"/>');
  });

  it('emits w:themeFontLang with bidi for Arabic', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ themeFontLang: { bidi: 'ar-SA' } });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/settings.xml');
    expect(xml).toContain('w:bidi="ar-SA"');
  });
});

// ---------------------------------------------------------------------------
// Page numbering (pgNumType)
// ---------------------------------------------------------------------------

describe('pgNumType in sectPr', () => {
  it('emits w:pgNumType with fmt', async () => {
    const b = new DocxBuilder();
    b.setPageSize({ w: 11906, h: 16838 });
    b.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });
    // Access sectPr via the builder's internal state through setPageSize
    // We need to set pgNumType — use the sectPr directly
    const b2 = new DocxBuilder();
    b2.addBlock(paragraph(['x']));
    // Override sectPr with pgNumType via the builder API
    // Since setPageSize/setPageMargins return this, we can chain
    // But pgNumType needs to be set — use the internal sectPr
    // The builder exposes setSectionType but not pgNumType directly yet
    // Test via the low-level approach: build a doc with pgNumType in pPr.sectPr
    b2.addBlock(paragraph(['section break'], {
      sectPr: {
        pgNumType: { fmt: 'upperRoman', start: 1 },
        pgSz: { w: 11906, h: 16838 },
        pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
      },
    }));
    const zip = await buildAndOpen(b2);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('w:fmt="upperRoman"');
    expect(xml).toContain('w:start="1"');
  });

  it('emits w:pgNumType with decimal fmt', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], {
      sectPr: { pgNumType: { fmt: 'decimal', start: 1 } },
    }));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('w:fmt="decimal"');
  });
});

// ---------------------------------------------------------------------------
// textDirection
// ---------------------------------------------------------------------------

describe('textDirection serialization', () => {
  it('emits w:textDirection on paragraph', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { textDirection: 'rl' }));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:textDirection w:val="rl"/>');
  });

  it('emits w:textDirection on sectPr', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], {
      sectPr: { textDirection: 'rl' },
    }));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:textDirection w:val="rl"/>');
  });
});

// ---------------------------------------------------------------------------
// Table cell completion
// ---------------------------------------------------------------------------

describe('table cell properties — Batch 1 additions', () => {
  it('emits w:tcBorders', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [{
        cells: [{
          tcPr: {
            tcBorders: {
              top:    { val: 'single', sz: 4, space: 0, color: 'FF0000' },
              bottom: { val: 'single', sz: 4, space: 0, color: '0000FF' },
            },
          },
          content: [paragraph(['x'])],
        }],
      }],
    });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:tcBorders>');
    expect(xml).toContain('w:color="FF0000"');
  });

  it('emits w:vMerge restart', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [
        { cells: [{ tcPr: { vMerge: 'restart' }, content: [paragraph(['top'])] }] },
        { cells: [{ tcPr: { vMerge: 'continue' }, content: [emptyParagraph()] }] },
      ],
    });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('<w:vMerge/>');
  });

  it('emits w:noWrap', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [{ cells: [{ tcPr: { noWrap: true }, content: [paragraph(['x'])] }] }],
    });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:noWrap/>');
  });

  it('emits w:tcMar', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [{
        cells: [{
          tcPr: {
            tcMar: {
              top:   { w: 100, type: 'dxa' },
              start: { w: 200, type: 'dxa' },
            },
          },
          content: [paragraph(['x'])],
        }],
      }],
    });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:tcMar>');
    expect(xml).toContain('w:w="100"');
    expect(xml).toContain('w:w="200"');
  });

  it('emits w:textDirection on cell', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [{ cells: [{ tcPr: { textDirection: 'rl' }, content: [paragraph(['x'])] }] }],
    });
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    expect(xml).toContain('<w:textDirection w:val="rl"/>');
  });
});

// ---------------------------------------------------------------------------
// Mid-document section break
// ---------------------------------------------------------------------------

describe('mid-document section break via pPr.sectPr', () => {
  it('emits w:sectPr inside w:pPr', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['section 1 end'], {
      sectPr: {
        type: 'nextPage',
        pgSz: { w: 11906, h: 16838 },
        pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
      },
    }));
    b.addBlock(paragraph(['section 2 start']));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/document.xml');
    // sectPr inside pPr (mid-doc break)
    expect(xml).toContain('<w:pPr>');
    expect(xml).toContain('<w:sectPr>');
    expect(xml).toContain('<w:type w:val="nextPage"/>');
  });
});

// ---------------------------------------------------------------------------
// Regression: existing tests still pass (no deferred parts accidentally emitted)
// ---------------------------------------------------------------------------

describe('Batch 1 regression — no accidental deferred parts', () => {
  const FORBIDDEN = [
    'word/header1.xml', 'word/footer1.xml',
    'word/footnotes.xml', 'word/endnotes.xml',
    'word/comments.xml', 'word/fontTable.xml',
  ];

  it('does not emit deferred parts when only numbering is used', async () => {
    const b = new DocxBuilder();
    b.defineNumbering({ levels: [decimalListLevel(0)] });
    b.addBlock(paragraph(['item'], { numPr: { numId: 1, ilvl: 0 } }));
    const zip = await buildAndOpen(b);
    const parts = partNames(zip);
    for (const p of FORBIDDEN) {
      expect(parts).not.toContain(p);
    }
  });

  it('does not emit deferred parts when settings are configured', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true, evenAndOddHeaders: true });
    b.addBlock(paragraph(['x']));
    const zip = await buildAndOpen(b);
    const parts = partNames(zip);
    for (const p of FORBIDDEN) {
      expect(parts).not.toContain(p);
    }
  });
});
