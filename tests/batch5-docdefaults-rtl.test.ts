/**
 * Batch 5 tests — Document Defaults and RTL/Arabic Support.
 *
 * Validates:
 * 1. w:docDefaults emitted in styles.xml when configureDocDefaults() is called
 * 2. w:rPrDefault with run properties (fonts, rtl, lang)
 * 3. w:pPrDefault with paragraph properties (bidi, jc, textDirection)
 * 4. No w:docDefaults emitted when not configured
 * 5. textDirection on sectPr (RTL section)
 * 6. textDirection on paragraph pPr
 * 7. RTL run properties (rtl, bCs, iCs, szCs, rFonts.cs)
 * 8. Arabic-friendly settings (bidi, themeFontLang)
 * 9. arabicAbjad numbering format
 * 10. Combined Arabic document: docDefaults + settings + RTL section
 * 11. No accidental deferred parts emitted
 */

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder, paragraph, textRun,
  ptToHalfPt,
} from '../src/index';
import type { WmlDocDefaults } from '../src/index';

async function buildAndOpen(b: DocxBuilder) {
  const buf = await b.build();
  return JSZip.loadAsync(buf);
}

async function readPart(zip: JSZip, name: string): Promise<string | null> {
  const entry = zip.file(name);
  if (!entry) return null;
  return entry.async('string');
}

// ---------------------------------------------------------------------------
// Document Defaults (docDefaults)
// ---------------------------------------------------------------------------

describe('Document Defaults (docDefaults)', () => {
  it('should NOT emit w:docDefaults when not configured', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['Hello']));
    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    expect(styles).not.toContain('w:docDefaults');
  });

  it('should emit w:docDefaults with rPrDefault only', async () => {
    const b = new DocxBuilder();
    b.configureDocDefaults({
      rPrDefault: { rFonts: { ascii: 'Times New Roman', cs: 'Arial Unicode MS' }, sz: ptToHalfPt(12) },
    });
    b.addBlock(paragraph(['Hello']));
    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    expect(styles).toContain('w:docDefaults');
    expect(styles).toContain('w:rPrDefault');
    expect(styles).toContain('Times New Roman');
    expect(styles).toContain('Arial Unicode MS');
    expect(styles).not.toContain('w:pPrDefault');
  });

  it('should emit w:docDefaults with pPrDefault only', async () => {
    const b = new DocxBuilder();
    b.configureDocDefaults({
      pPrDefault: { bidi: true, jc: 'end' },
    });
    b.addBlock(paragraph(['Hello']));
    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    expect(styles).toContain('w:docDefaults');
    expect(styles).toContain('w:pPrDefault');
    expect(styles).toContain('<w:bidi/>');
    expect(styles).toContain('w:jc');
    expect(styles).not.toContain('w:rPrDefault');
  });

  it('should emit w:docDefaults with both rPrDefault and pPrDefault', async () => {
    const b = new DocxBuilder();
    const docDefaults: WmlDocDefaults = {
      rPrDefault: {
        rtl: true,
        rFonts: { cs: 'Arial Unicode MS', ascii: 'Times New Roman' },
        lang: { bidi: 'ar-SA', val: 'en-US' },
      },
      pPrDefault: {
        bidi: true,
        jc: 'end',
      },
    };
    b.configureDocDefaults(docDefaults);
    b.addBlock(paragraph(['Hello']));
    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    expect(styles).toContain('w:docDefaults');
    expect(styles).toContain('w:rPrDefault');
    expect(styles).toContain('w:pPrDefault');
    expect(styles).toContain('<w:rtl/>');
    expect(styles).toContain('w:bidi="ar-SA"');
    expect(styles).toContain('w:val="en-US"');
  });
});

// ---------------------------------------------------------------------------
// Text Direction
// ---------------------------------------------------------------------------

describe('Text Direction', () => {
  it('should serialize textDirection in sectPr', async () => {
    const b = new DocxBuilder();
    b.setSectionTextDirection('rl');
    b.addBlock(paragraph(['RTL section']));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');
    expect(doc).toContain('<w:textDirection w:val="rl"/>');
  });

  it('should serialize textDirection in paragraph pPr', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['RTL paragraph'], { textDirection: 'rl', bidi: true }));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');
    expect(doc).toContain('w:textDirection');
    expect(doc).toContain('w:val="rl"');
    expect(doc).toContain('<w:bidi/>');
  });
});

// ---------------------------------------------------------------------------
// RTL and Arabic Support
// ---------------------------------------------------------------------------

describe('RTL and Arabic Support', () => {
  it('should serialize RTL paragraph properties', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(
      [{ text: 'مرحبا', rPr: { rtl: true, cs: true, rFonts: { cs: 'Arial Unicode MS' } } }],
      { bidi: true, jc: 'end' },
    ));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');
    expect(doc).toContain('<w:bidi/>');
    expect(doc).toContain('w:jc w:val="end"');
    expect(doc).toContain('<w:rtl/>');
    expect(doc).toContain('<w:cs/>');
    expect(doc).toContain('w:cs="Arial Unicode MS"');
  });

  it('should serialize RTL run properties', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([
      textRun('Arabic text', {
        rtl: true,
        bCs: true,
        szCs: ptToHalfPt(14),
        rFonts: { cs: 'Traditional Arabic', ascii: 'Times New Roman' },
        lang: { bidi: 'ar-SA' },
      }),
    ]));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');
    expect(doc).toContain('<w:rtl/>');
    expect(doc).toContain('<w:bCs/>');
    expect(doc).toContain('w:szCs w:val="28"');
    expect(doc).toContain('w:cs="Traditional Arabic"');
    expect(doc).toContain('w:bidi="ar-SA"');
  });

  it('should support arabicAbjad numbering', async () => {
    const b = new DocxBuilder();
    const numId = b.defineNumbering({
      levels: [{
        ilvl: 0,
        numFmt: 'arabicAbjad',
        lvlText: '%1.',
        lvlJc: 'start',
        suff: 'tab',
        pPr: { bidi: true, ind: { start: 720, hanging: 360 } },
      }],
    });
    b.addBlock(paragraph(['بند أول'], { numPr: { numId, ilvl: 0 }, bidi: true }));
    const zip = await buildAndOpen(b);
    const numbering = await readPart(zip, 'word/numbering.xml');
    expect(numbering).toContain('arabicAbjad');
    const doc = await readPart(zip, 'word/document.xml');
    expect(doc).toContain('w:numId');
  });

  it('should serialize themeFontLang with bidi support in settings', async () => {
    const b = new DocxBuilder();
    b.configureSettings({
      themeFontLang: { bidi: 'ar-SA', val: 'en-US' },
      updateFields: true,
    });
    b.addBlock(paragraph(['Test']));
    const zip = await buildAndOpen(b);
    const settings = await readPart(zip, 'word/settings.xml');
    expect(settings).toContain('w:themeFontLang');
    expect(settings).toContain('w:bidi="ar-SA"');
    expect(settings).toContain('w:val="en-US"');
  });

  it('should create complete Arabic document with docDefaults', async () => {
    const b = new DocxBuilder();

    // Arabic document defaults
    b.configureDocDefaults({
      rPrDefault: {
        rtl: true,
        rFonts: { cs: 'Traditional Arabic', ascii: 'Times New Roman', hAnsi: 'Times New Roman' },
        lang: { bidi: 'ar-SA', val: 'en-US' },
        szCs: ptToHalfPt(14),
      },
      pPrDefault: {
        bidi: true,
        jc: 'end',
      },
    });

    // RTL settings
    b.configureSettings({
      bidi: true,
      themeFontLang: { bidi: 'ar-SA', val: 'en-US' },
      updateFields: true,
    });

    // RTL section
    b.setSectionProperties({
      pgSz: { w: 11906, h: 16838 },
      pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
      bidi: true,
      rtlGutter: true,
    });

    // Arabic content
    b.addBlock(paragraph(
      [{ text: 'مرحبا بالعالم', rPr: { rtl: true, cs: true } }],
      { bidi: true, jc: 'end' },
    ));
    b.addBlock(paragraph(
      [{ text: 'Hello World', rPr: { rtl: false } }],
      { jc: 'start' },
    ));

    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    const doc = await readPart(zip, 'word/document.xml');
    const settings = await readPart(zip, 'word/settings.xml');

    // docDefaults in styles.xml
    expect(styles).toContain('w:docDefaults');
    expect(styles).toContain('Traditional Arabic');
    expect(styles).toContain('w:bidi="ar-SA"');

    // RTL section in document.xml
    expect(doc).toContain('<w:bidi/>');
    expect(doc).toContain('<w:rtlGutter/>');

    // settings.xml
    expect(settings).toContain('w:bidi="ar-SA"');
  });
});

// ---------------------------------------------------------------------------
// Regression: no accidental deferred parts
// ---------------------------------------------------------------------------

describe('Batch 5 regression — no accidental deferred parts', () => {
  it('does not emit fontTable.xml or comments.xml with docDefaults', async () => {
    const b = new DocxBuilder();
    b.configureDocDefaults({
      rPrDefault: { rtl: true, rFonts: { cs: 'Arial Unicode MS' } },
      pPrDefault: { bidi: true },
    });
    b.addBlock(paragraph(['Test']));
    const zip = await buildAndOpen(b);
    const parts = Object.keys(zip.files).filter(n => !zip.files[n].dir);
    expect(parts).not.toContain('word/fontTable.xml');
    expect(parts).not.toContain('word/comments.xml');
    expect(parts).not.toContain('word/theme/theme1.xml');
  });

  it('styles.xml is valid XML with docDefaults', async () => {
    const b = new DocxBuilder();
    b.configureDocDefaults({
      rPrDefault: { rtl: true, rFonts: { cs: 'Traditional Arabic' }, lang: { bidi: 'ar-SA' } },
      pPrDefault: { bidi: true, jc: 'end' },
    });
    b.addBlock(paragraph(['Test']));
    const zip = await buildAndOpen(b);
    const styles = await readPart(zip, 'word/styles.xml');
    // Basic well-formedness checks
    expect(styles).toMatch(/^<\?xml/);
    expect(styles).toContain('<w:styles');
    expect(styles).toContain('</w:styles>');
    expect(styles).not.toContain('w::');
  });
});
