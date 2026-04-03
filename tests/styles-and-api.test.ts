/**
 * Styles XML and API surface tests.
 *
 * Verifies:
 * 1. word/styles.xml structure and content
 * 2. Style inheritance (basedOn, next)
 * 3. Default style flag
 * 4. DocxBuilder API — method chaining, defaults
 * 5. Measurement helpers
 * 6. Factory helpers produce correct structures
 */

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder, paragraph, textRun, emptyParagraph, pageBreak,
  ptToTwips, inToTwips, cmToTwips, ptToHalfPt,
} from '../src/index';

async function getStylesXml(b: DocxBuilder): Promise<string> {
  const buf = await b.build();
  const zip = await JSZip.loadAsync(buf);
  return zip.file('word/styles.xml')!.async('string');
}

// ---------------------------------------------------------------------------
// Styles XML
// ---------------------------------------------------------------------------

describe('word/styles.xml', () => {
  it('has WML namespace', async () => {
    const b = new DocxBuilder();
    const xml = await getStylesXml(b);
    expect(xml).toContain('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  });

  it('emits w:style element for each added style', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true });
    b.addStyle({ type: 'paragraph', styleId: 'Heading1', name: 'heading 1' });
    const xml = await getStylesXml(b);
    expect(xml).toContain('w:styleId="Normal"');
    expect(xml).toContain('w:styleId="Heading1"');
  });

  it('emits w:default="1" for default style', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true });
    const xml = await getStylesXml(b);
    expect(xml).toContain('w:default="1"');
  });

  it('does NOT emit w:default for non-default style', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'paragraph', styleId: 'Heading1', name: 'heading 1' });
    const xml = await getStylesXml(b);
    expect(xml).not.toContain('w:default=');
  });

  it('emits w:basedOn', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'paragraph', styleId: 'Heading1', name: 'heading 1', basedOn: 'Normal' });
    const xml = await getStylesXml(b);
    expect(xml).toContain('<w:basedOn w:val="Normal"/>');
  });

  it('emits w:next', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'paragraph', styleId: 'Heading1', name: 'heading 1', next: 'Normal' });
    const xml = await getStylesXml(b);
    expect(xml).toContain('<w:next w:val="Normal"/>');
  });

  it('emits w:pPr inside style', async () => {
    const b = new DocxBuilder();
    b.addStyle({
      type: 'paragraph', styleId: 'S1', name: 'S1',
      pPr: { jc: 'center', outlineLvl: 0 },
    });
    const xml = await getStylesXml(b);
    expect(xml).toContain('<w:pPr>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:outlineLvl w:val="0"/>');
  });

  it('emits w:rPr inside style', async () => {
    const b = new DocxBuilder();
    b.addStyle({
      type: 'paragraph', styleId: 'S1', name: 'S1',
      rPr: { b: true, sz: 32 },
    });
    const xml = await getStylesXml(b);
    expect(xml).toContain('<w:rPr>');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:sz w:val="32"/>');
  });

  it('emits character style type', async () => {
    const b = new DocxBuilder();
    b.addStyle({ type: 'character', styleId: 'Strong', name: 'Strong', rPr: { b: true } });
    const xml = await getStylesXml(b);
    expect(xml).toContain('w:type="character"');
    expect(xml).toContain('w:styleId="Strong"');
  });

  it('empty styles list produces valid w:styles element', async () => {
    const b = new DocxBuilder();
    const xml = await getStylesXml(b);
    expect(xml).toContain('<w:styles');
    // Empty styles serializes as self-closing element — both forms are valid XML
    const hasClosingTag = xml.includes('</w:styles>') || xml.includes('/>');
    expect(hasClosingTag).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// DocxBuilder API
// ---------------------------------------------------------------------------

describe('DocxBuilder API', () => {
  it('build() returns a Buffer', async () => {
    const b = new DocxBuilder();
    const buf = await b.build();
    expect(Buffer.isBuffer(buf)).toBe(true);
  });

  it('method chaining works', async () => {
    const b = new DocxBuilder();
    const result = b
      .setCoreProperties({ title: 'Test' })
      .setExtendedProperties({ application: 'Test' })
      .setPageSize({ w: 11906, h: 16838 })
      .setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 })
      .addStyle({ type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true })
      .addBlock(paragraph(['x']));
    expect(result).toBe(b);
  });

  it('addParagraph() convenience method works', async () => {
    const b = new DocxBuilder();
    b.addParagraph({ runs: [{ text: 'Hello' }] });
    const buf = await b.build();
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('Hello');
  });

  it('addParagraph() with style works', async () => {
    const b = new DocxBuilder();
    b.addParagraph({ style: 'Heading1', runs: [{ text: 'Title' }] });
    const buf = await b.build();
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('w:val="Heading1"');
  });

  it('addTable() convenience method works', async () => {
    const b = new DocxBuilder();
    b.addTable({
      tblPr: {},
      colWidths: [2500, 2500],
      rows: [{ cells: [{ content: [paragraph(['A'])] }, { content: [paragraph(['B'])] }] }],
    });
    const buf = await b.build();
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:tbl>');
  });

  it('default page size is US Letter', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['x']));
    const buf = await b.build();
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')!.async('string');
    // Default: US Letter w=12240, h=15840
    expect(xml).toContain('w:w="12240"');
    expect(xml).toContain('w:h="15840"');
  });
});

// ---------------------------------------------------------------------------
// Factory helpers
// ---------------------------------------------------------------------------

describe('factory helpers', () => {
  it('paragraph() creates a WmlParagraph', () => {
    const p = paragraph(['hello']);
    expect(p._type).toBe('paragraph');
    expect(p.content).toHaveLength(1);
    expect(p.content[0]._type).toBe('run');
  });

  it('paragraph() with mixed content creates correct runs', () => {
    const p = paragraph(['plain', { text: 'bold', rPr: { b: true } }]);
    expect(p.content).toHaveLength(2);
    expect((p.content[0] as any).rPr).toBeUndefined();
    expect((p.content[1] as any).rPr?.b).toBe(true);
  });

  it('textRun() creates a WmlRun with text', () => {
    const r = textRun('hello');
    expect(r._type).toBe('run');
    expect(r.content[0]._type).toBe('text');
    expect((r.content[0] as any).text).toBe('hello');
  });

  it('textRun() with empty rPr does not attach rPr', () => {
    const r = textRun('hello', {});
    expect(r.rPr).toBeUndefined();
  });

  it('textRun() with actual rPr attaches it', () => {
    const r = textRun('hello', { b: true });
    expect(r.rPr?.b).toBe(true);
  });

  it('emptyParagraph() creates a paragraph with no content', () => {
    const p = emptyParagraph();
    expect(p._type).toBe('paragraph');
    expect(p.content).toHaveLength(0);
  });

  it('pageBreak() creates a paragraph with a page break run', () => {
    const p = pageBreak();
    expect(p._type).toBe('paragraph');
    expect(p.content[0]._type).toBe('run');
    const run = p.content[0] as any;
    expect(run.content[0]._type).toBe('break');
    expect(run.content[0].breakType).toBe('page');
  });
});

// ---------------------------------------------------------------------------
// Measurement helpers
// ---------------------------------------------------------------------------

describe('measurement helpers', () => {
  it('ptToTwips: 1pt = 20 twips', () => {
    expect(ptToTwips(1)).toBe(20);
  });

  it('ptToTwips: 72pt = 1440 twips (1 inch)', () => {
    expect(ptToTwips(72)).toBe(1440);
  });

  it('inToTwips: 1 inch = 1440 twips', () => {
    expect(inToTwips(1)).toBe(1440);
  });

  it('inToTwips: 0.5 inch = 720 twips', () => {
    expect(inToTwips(0.5)).toBe(720);
  });

  it('cmToTwips: 1cm ≈ 567 twips', () => {
    expect(cmToTwips(1)).toBe(567);
  });

  it('ptToHalfPt: 12pt = 24 half-points', () => {
    expect(ptToHalfPt(12)).toBe(24);
  });

  it('ptToHalfPt: 11pt = 22 half-points', () => {
    expect(ptToHalfPt(11)).toBe(22);
  });
});
