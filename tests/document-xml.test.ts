/**
 * word/document.xml content tests.
 *
 * Verifies:
 * 1. No empty w:rPr/ elements
 * 2. xml:space="preserve" on w:t with whitespace
 * 3. Correct WML namespace
 * 4. w:conformance="strict"
 * 5. w:sectPr is last child of w:body
 * 6. Paragraph formatting (alignment, spacing, indentation)
 * 7. Run formatting (bold, italic, color, font size, underline)
 * 8. Page break
 * 9. Table structure (tblPr, tblGrid, tr, tc)
 * 10. Table cell must contain at least one block element
 */

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder, paragraph, textRun, emptyParagraph, pageBreak,
  ptToTwips, ptToHalfPt, inToTwips,
} from '../src/index';

async function getDocXml(b: DocxBuilder): Promise<string> {
  const buf = await b.build();
  const zip = await JSZip.loadAsync(buf);
  return zip.file('word/document.xml')!.async('string');
}

// ---------------------------------------------------------------------------
// Namespace and conformance
// ---------------------------------------------------------------------------

describe('document.xml — namespace and conformance', () => {
  it('uses WML strict namespace', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toContain('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  });

  it('does not emit w:conformance (Word rejects it)', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).not.toContain('w:conformance');
  });

  it('has w:document as root element', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toMatch(/<w:document\b/);
    expect(xml).toContain('</w:document>');
  });
});

// ---------------------------------------------------------------------------
// Empty w:rPr guard
// ---------------------------------------------------------------------------

describe('document.xml — no empty w:rPr/', () => {
  it('plain text run has no w:rPr/', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['plain text']));
    const xml = await getDocXml(b);
    expect(xml).not.toContain('<w:rPr/>');
  });

  it('run with empty rPr object has no w:rPr/', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'text', rPr: {} }]));
    const xml = await getDocXml(b);
    expect(xml).not.toContain('<w:rPr/>');
  });

  it('run with actual rPr properties emits w:rPr', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'bold', rPr: { b: true } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:rPr>');
    expect(xml).toContain('<w:b/>');
  });

  it('mixed paragraph with some plain runs has no empty w:rPr/', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([
      'plain',
      { text: 'bold', rPr: { b: true } },
      'plain again',
    ]));
    const xml = await getDocXml(b);
    expect(xml).not.toContain('<w:rPr/>');
  });
});

// ---------------------------------------------------------------------------
// xml:space="preserve"
// ---------------------------------------------------------------------------

describe('document.xml — xml:space="preserve"', () => {
  it('sets xml:space="preserve" on text with leading space', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([' leading space']));
    const xml = await getDocXml(b);
    expect(xml).toContain('xml:space="preserve"');
  });

  it('sets xml:space="preserve" on text with trailing space', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['trailing space ']));
    const xml = await getDocXml(b);
    expect(xml).toContain('xml:space="preserve"');
  });

  it('sets xml:space="preserve" on all text runs (always safe)', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['no spaces']));
    const xml = await getDocXml(b);
    // textRun always sets space: 'preserve' — safe default
    expect(xml).toContain('xml:space="preserve"');
  });
});

// ---------------------------------------------------------------------------
// Section properties
// ---------------------------------------------------------------------------

describe('document.xml — section properties', () => {
  it('w:sectPr is present', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:sectPr>');
  });

  it('w:sectPr is the last element before </w:body>', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    const bodyClose = xml.indexOf('</w:body>');
    const sectPrClose = xml.lastIndexOf('</w:sectPr>');
    expect(sectPrClose).toBeLessThan(bodyClose);
    expect(sectPrClose).toBeGreaterThan(0);
    // Nothing between </w:sectPr> and </w:body>
    const between = xml.slice(sectPrClose + '</w:sectPr>'.length, bodyClose).trim();
    expect(between).toBe('');
  });

  it('emits w:pgSz with A4 dimensions', async () => {
    const b = new DocxBuilder();
    b.setPageSize({ w: 11906, h: 16838 });
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:w="11906"');
    expect(xml).toContain('w:h="16838"');
  });

  it('emits w:pgSz with landscape orientation', async () => {
    const b = new DocxBuilder();
    b.setPageSize({ w: 16838, h: 11906, orient: 'landscape' });
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:orient="landscape"');
  });

  it('emits w:pgMar with all required attributes', async () => {
    const b = new DocxBuilder();
    b.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });
    b.addBlock(paragraph(['x']));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:top="1440"');
    expect(xml).toContain('w:right="1800"');
    expect(xml).toContain('w:bottom="1440"');
    expect(xml).toContain('w:left="1800"');
    expect(xml).toContain('w:header="720"');
    expect(xml).toContain('w:footer="720"');
    expect(xml).toContain('w:gutter="0"');
  });
});

// ---------------------------------------------------------------------------
// Paragraph formatting
// ---------------------------------------------------------------------------

describe('document.xml — paragraph formatting', () => {
  it('emits w:pStyle', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { pStyle: 'Heading1' }));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>');
  });

  it('emits w:jc for alignment', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { jc: 'center' }));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:jc w:val="center"/>');
  });

  it('emits w:spacing', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { spacing: { before: 240, after: 120 } }));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:before="240"');
    expect(xml).toContain('w:after="120"');
  });

  it('emits w:ind with firstLine', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { ind: { firstLine: inToTwips(0.5) } }));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:firstLine="720"');
  });

  it('emits w:outlineLvl', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { outlineLvl: 0 }));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:outlineLvl w:val="0"/>');
  });

  it('emits w:keepNext', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { keepNext: true }));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:keepNext/>');
  });

  it('emits w:keepNext with val="0" when false', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x'], { keepNext: false }));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:keepNext w:val="0"/>');
  });
});

// ---------------------------------------------------------------------------
// Run formatting
// ---------------------------------------------------------------------------

describe('document.xml — run formatting', () => {
  it('emits w:b for bold', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { b: true } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:b/>');
  });

  it('emits w:i for italic', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { i: true } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:i/>');
  });

  it('emits w:u for underline', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { u: 'single' } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:u w:val="single"/>');
  });

  it('emits w:strike for strikethrough', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { strike: true } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:strike/>');
  });

  it('emits w:color', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { color: { val: 'FF0000' } } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:color w:val="FF0000"/>');
  });

  it('emits w:sz for font size', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { sz: ptToHalfPt(12) } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:sz w:val="24"/>');
  });

  it('emits w:rFonts for font family', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { rFonts: { ascii: 'Arial', hAnsi: 'Arial' } } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:ascii="Arial"');
    expect(xml).toContain('w:hAnsi="Arial"');
  });

  it('emits w:vertAlign for superscript', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { vertAlign: 'superscript' } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:vertAlign w:val="superscript"/>');
  });

  it('emits w:vanish for hidden text', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { vanish: true } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:vanish/>');
  });

  it('emits w:highlight', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { highlight: 'yellow' } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:highlight w:val="yellow"/>');
  });

  it('emits w:lang', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([{ text: 'x', rPr: { lang: { val: 'en-US' } } }]));
    const xml = await getDocXml(b);
    expect(xml).toContain('w:val="en-US"');
  });
});

// ---------------------------------------------------------------------------
// Page break
// ---------------------------------------------------------------------------

describe('document.xml — page break', () => {
  it('emits w:br with type="page"', async () => {
    const b = new DocxBuilder();
    b.addBlock(pageBreak());
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:br w:type="page"/>');
  });
});

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

describe('document.xml — tables', () => {
  function makeTable(b: DocxBuilder) {
    b.addTable({
      tblPr: { tblW: { w: 5000, type: 'dxa' } },
      colWidths: [2500, 2500],
      rows: [
        { cells: [
          { content: [paragraph(['A'])] },
          { content: [paragraph(['B'])] },
        ]},
        { cells: [
          { content: [paragraph(['C'])] },
          { content: [paragraph(['D'])] },
        ]},
      ],
    });
  }

  it('emits w:tbl', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:tbl>');
  });

  it('emits w:tblPr before w:tblGrid', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    expect(xml.indexOf('<w:tblPr>')).toBeLessThan(xml.indexOf('<w:tblGrid>'));
  });

  it('emits w:tblGrid with correct column count', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    const gridCols = [...xml.matchAll(/<w:gridCol\b/g)];
    expect(gridCols.length).toBe(2);
  });

  it('emits correct column widths', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    expect(xml).toContain('w:w="2500"');
  });

  it('emits w:tr for each row', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    const rows = [...xml.matchAll(/<w:tr>/g)];
    expect(rows.length).toBe(2);
  });

  it('emits w:tc for each cell', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    const cells = [...xml.matchAll(/<w:tc>/g)];
    expect(cells.length).toBe(4);
  });

  it('each cell contains at least one w:p', async () => {
    const b = new DocxBuilder();
    makeTable(b);
    const xml = await getDocXml(b);
    // Each tc must have a p inside it
    const tcBlocks = xml.split('<w:tc>').slice(1);
    for (const block of tcBlocks) {
      const tcContent = block.slice(0, block.indexOf('</w:tc>'));
      expect(tcContent).toContain('<w:p>');
    }
  });

  it('emits table borders when specified', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {
        tblBorders: {
          top:    { val: 'single', sz: 4, space: 0, color: '000000' },
          bottom: { val: 'single', sz: 4, space: 0, color: '000000' },
          start:  { val: 'single', sz: 4, space: 0, color: '000000' },
          end:    { val: 'single', sz: 4, space: 0, color: '000000' },
        },
      },
      colWidths: [2500],
      rows: [{ cells: [{ content: [paragraph(['x'])] }] }],
    });
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:tblBorders>');
    expect(xml).toContain('w:val="single"');
  });

  it('emits gridSpan for merged cells', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2000, 2000, 2000],
      rows: [{
        cells: [
          { tcPr: { gridSpan: 2 }, content: [paragraph(['merged'])] },
          { content: [paragraph(['single'])] },
        ],
      }],
    });
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
  });

  it('empty cell gets an empty paragraph placeholder', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500],
      rows: [{ cells: [{ content: [] }] }],  // empty content
    });
    const xml = await getDocXml(b);
    // Builder should insert emptyParagraph() for empty cells
    expect(xml).toContain('<w:tc>');
    expect(xml).toContain('<w:p');
  });
});

// ---------------------------------------------------------------------------
// Text content
// ---------------------------------------------------------------------------

describe('document.xml — text content', () => {
  it('text content is correctly escaped', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['<test> & "quotes"']));
    const xml = await getDocXml(b);
    // XML text nodes require escaping of <, >, & — quotes are valid unescaped in text content
    expect(xml).toContain('&lt;test&gt; &amp; "quotes"');
  });

  it('empty paragraph produces w:p with no children', async () => {
    const b = new DocxBuilder();
    b.addBlock(emptyParagraph());
    const xml = await getDocXml(b);
    expect(xml).toContain('<w:p/>');
  });

  it('multiple paragraphs are all present', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['First']));
    b.addBlock(paragraph(['Second']));
    b.addBlock(paragraph(['Third']));
    const xml = await getDocXml(b);
    expect(xml).toContain('First');
    expect(xml).toContain('Second');
    expect(xml).toContain('Third');
  });
});
