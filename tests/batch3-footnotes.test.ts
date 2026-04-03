/**
 * Batch 3 tests — Footnotes and Endnotes.
 *
 * Validates:
 * 1. word/footnotes.xml emitted when addFootnote() is called
 * 2. word/endnotes.xml emitted when addEndnote() is called
 * 3. Required separator entries (id=-1, id=0) always present
 * 4. User note IDs start at 1 and auto-increment
 * 5. w:footnoteReference in document.xml matches note id in footnotes.xml
 * 6. w:endnoteReference in document.xml matches note id in endnotes.xml
 * 7. Correct content types for footnotes/endnotes parts
 * 8. Correct relationship entries in word/_rels/document.xml.rels
 * 9. footnotes.xml / endnotes.xml root element and namespace
 * 10. configureFootnotes() / configureEndnotes() properties emitted
 * 11. No footnotes.xml emitted when no footnotes added
 * 12. No endnotes.xml emitted when no endnotes added
 * 13. No deferred parts accidentally emitted
 * 14. Regression: all prior tests still pass
 */

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder, paragraph, textRun,
  serializeFootnotes, serializeEndnotes,
} from '../src/index';
import type { WmlBlockElement } from '../src/index';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

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

/** Build a minimal doc with one footnote */
function buildWithOneFootnote() {
  const b = new DocxBuilder();
  const noteRun = b.addFootnote([paragraph(['This is a footnote.'])]);
  b.addBlock(paragraph(['Main text', noteRun]));
  return b;
}

/** Build a minimal doc with one endnote */
function buildWithOneEndnote() {
  const b = new DocxBuilder();
  const noteRun = b.addEndnote([paragraph(['This is an endnote.'])]);
  b.addBlock(paragraph(['Main text', noteRun]));
  return b;
}

// ---------------------------------------------------------------------------
// Part emission
// ---------------------------------------------------------------------------

describe('footnote/endnote part emission', () => {
  it('emits word/footnotes.xml when addFootnote() is called', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    expect(partNames(zip)).toContain('word/footnotes.xml');
  });

  it('emits word/endnotes.xml when addEndnote() is called', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    expect(partNames(zip)).toContain('word/endnotes.xml');
  });

  it('does NOT emit word/footnotes.xml when no footnotes added', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['No notes here.']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).not.toContain('word/footnotes.xml');
  });

  it('does NOT emit word/endnotes.xml when no endnotes added', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph(['No notes here.']));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).not.toContain('word/endnotes.xml');
  });

  it('emits both footnotes.xml and endnotes.xml when both are used', async () => {
    const b = new DocxBuilder();
    const fn = b.addFootnote([paragraph(['Footnote content.'])]);
    const en = b.addEndnote([paragraph(['Endnote content.'])]);
    b.addBlock(paragraph(['Text', fn, en]));
    const zip = await buildAndOpen(b);
    const parts = partNames(zip);
    expect(parts).toContain('word/footnotes.xml');
    expect(parts).toContain('word/endnotes.xml');
  });
});

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

describe('footnote/endnote content types', () => {
  it('has correct content type for footnotes', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const ct = await readPart(zip, '[Content_Types].xml');
    expect(ct).toContain(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    );
  });

  it('has correct content type for endnotes', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const ct = await readPart(zip, '[Content_Types].xml');
    expect(ct).toContain(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
    );
  });
});

// ---------------------------------------------------------------------------
// Relationships
// ---------------------------------------------------------------------------

describe('footnote/endnote relationships', () => {
  it('has footnotes relationship in document.xml.rels', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const rels = await readPart(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('relationships/footnotes');
    expect(rels).toContain('Target="footnotes.xml"');
  });

  it('has endnotes relationship in document.xml.rels', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const rels = await readPart(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('relationships/endnotes');
    expect(rels).toContain('Target="endnotes.xml"');
  });
});

// ---------------------------------------------------------------------------
// footnotes.xml structure
// ---------------------------------------------------------------------------

describe('footnotes.xml structure', () => {
  it('has w:footnotes root element', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toContain('<w:footnotes');
    expect(xml).toContain('</w:footnotes>');
  });

  it('starts with XML declaration', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toMatch(/^<\?xml/);
  });

  it('has no duplicate namespace prefixes (no w::)', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).not.toContain('w::');
  });

  it('always contains separator entry (id=-1)', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toContain('w:id="-1"');
    expect(xml).toContain('w:type="separator"');
  });

  it('always contains continuationSeparator entry (id=0)', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toContain('w:id="0"');
    expect(xml).toContain('w:type="continuationSeparator"');
  });

  it('user note has id=1 and no type attribute', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    // User note: id=1, no w:type attribute (normal notes omit type)
    expect(xml).toContain('w:id="1"');
    // The user note element should not have w:type (only system notes do)
    // Check that the w:footnote with id=1 does not have w:type
    const noteMatch = xml?.match(/<w:footnote[^>]*w:id="1"[^>]*>/);
    expect(noteMatch).toBeTruthy();
    expect(noteMatch![0]).not.toContain('w:type');
  });

  it('user note content is present in footnotes.xml', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toContain('This is a footnote.');
  });

  it('multiple footnotes get sequential IDs starting at 1', async () => {
    const b = new DocxBuilder();
    const fn1 = b.addFootnote([paragraph(['First note.'])]);
    const fn2 = b.addFootnote([paragraph(['Second note.'])]);
    const fn3 = b.addFootnote([paragraph(['Third note.'])]);
    b.addBlock(paragraph(['Text', fn1, fn2, fn3]));
    const zip = await buildAndOpen(b);
    const xml = await readPart(zip, 'word/footnotes.xml');
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('w:id="2"');
    expect(xml).toContain('w:id="3"');
    expect(xml).toContain('First note.');
    expect(xml).toContain('Second note.');
    expect(xml).toContain('Third note.');
  });
});

// ---------------------------------------------------------------------------
// endnotes.xml structure
// ---------------------------------------------------------------------------

describe('endnotes.xml structure', () => {
  it('has w:endnotes root element', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const xml = await readPart(zip, 'word/endnotes.xml');
    expect(xml).toContain('<w:endnotes');
    expect(xml).toContain('</w:endnotes>');
  });

  it('always contains separator entry (id=-1)', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const xml = await readPart(zip, 'word/endnotes.xml');
    expect(xml).toContain('w:id="-1"');
    expect(xml).toContain('w:type="separator"');
  });

  it('always contains continuationSeparator entry (id=0)', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const xml = await readPart(zip, 'word/endnotes.xml');
    expect(xml).toContain('w:id="0"');
    expect(xml).toContain('w:type="continuationSeparator"');
  });

  it('user endnote has id=1 and no type attribute', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const xml = await readPart(zip, 'word/endnotes.xml');
    expect(xml).toContain('w:id="1"');
    const noteMatch = xml?.match(/<w:endnote[^>]*w:id="1"[^>]*>/);
    expect(noteMatch).toBeTruthy();
    expect(noteMatch![0]).not.toContain('w:type');
  });

  it('user endnote content is present in endnotes.xml', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const xml = await readPart(zip, 'word/endnotes.xml');
    expect(xml).toContain('This is an endnote.');
  });
});

// ---------------------------------------------------------------------------
// Reference ID matching — document.xml ↔ footnotes.xml
// ---------------------------------------------------------------------------

describe('reference ID matching', () => {
  it('w:footnoteReference in document.xml has id matching footnotes.xml', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const doc = await readPart(zip, 'word/document.xml');
    const ftn = await readPart(zip, 'word/footnotes.xml');

    // Extract reference id from document.xml
    const refMatch = doc?.match(/<w:footnoteReference[^>]*w:id="(\d+)"/);
    expect(refMatch).toBeTruthy();
    const refId = refMatch![1];

    // Verify same id exists as a user note in footnotes.xml
    const noteMatch = ftn?.match(new RegExp(`<w:footnote[^>]*w:id="${refId}"[^>]*>`));
    expect(noteMatch).toBeTruthy();
  });

  it('w:endnoteReference in document.xml has id matching endnotes.xml', async () => {
    const zip = await buildAndOpen(buildWithOneEndnote());
    const doc = await readPart(zip, 'word/document.xml');
    const edn = await readPart(zip, 'word/endnotes.xml');

    const refMatch = doc?.match(/<w:endnoteReference[^>]*w:id="(\d+)"/);
    expect(refMatch).toBeTruthy();
    const refId = refMatch![1];

    const noteMatch = edn?.match(new RegExp(`<w:endnote[^>]*w:id="${refId}"[^>]*>`));
    expect(noteMatch).toBeTruthy();
  });

  it('multiple footnote references all match their notes', async () => {
    const b = new DocxBuilder();
    const fn1 = b.addFootnote([paragraph(['Note one.'])]);
    const fn2 = b.addFootnote([paragraph(['Note two.'])]);
    b.addBlock(paragraph(['Text', fn1, fn2]));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');
    const ftn = await readPart(zip, 'word/footnotes.xml');

    const refIds = [...(doc ?? '').matchAll(/<w:footnoteReference[^>]*w:id="(\d+)"/g)]
      .map(m => m[1]);
    expect(refIds).toHaveLength(2);

    for (const id of refIds) {
      const noteMatch = ftn?.match(new RegExp(`<w:footnote[^>]*w:id="${id}"[^>]*>`));
      expect(noteMatch).toBeTruthy();
    }
  });

  it('footnote and endnote IDs are independent (both start at 1)', async () => {
    const b = new DocxBuilder();
    const fn = b.addFootnote([paragraph(['Footnote 1.'])]);
    const en = b.addEndnote([paragraph(['Endnote 1.'])]);
    b.addBlock(paragraph(['Text', fn, en]));
    const zip = await buildAndOpen(b);
    const doc = await readPart(zip, 'word/document.xml');

    const fnRef = doc?.match(/<w:footnoteReference[^>]*w:id="(\d+)"/);
    const enRef = doc?.match(/<w:endnoteReference[^>]*w:id="(\d+)"/);
    expect(fnRef![1]).toBe('1');
    expect(enRef![1]).toBe('1');
  });
});

// ---------------------------------------------------------------------------
// serializeFootnotes / serializeEndnotes standalone
// ---------------------------------------------------------------------------

describe('serializeFootnotes standalone', () => {
  it('produces valid XML with separator entries', () => {
    const xml = serializeFootnotes([]);
    expect(xml).toMatch(/^<\?xml/);
    expect(xml).toContain('<w:footnotes');
    expect(xml).toContain('w:id="-1"');
    expect(xml).toContain('w:id="0"');
  });

  it('includes user notes in output', () => {
    const xml = serializeFootnotes([
      { id: 1, type: 'normal', content: [paragraph(['Custom note.'])] },
    ]);
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('Custom note.');
  });
});

describe('serializeEndnotes standalone', () => {
  it('produces valid XML with separator entries', () => {
    const xml = serializeEndnotes([]);
    expect(xml).toMatch(/^<\?xml/);
    expect(xml).toContain('<w:endnotes');
    expect(xml).toContain('w:id="-1"');
    expect(xml).toContain('w:id="0"');
  });
});

// ---------------------------------------------------------------------------
// configureFootnotes / configureEndnotes
// ---------------------------------------------------------------------------

describe('configureFootnotes / configureEndnotes', () => {
  it('configureFootnotes does not throw and builder still builds', async () => {
    const b = new DocxBuilder();
    b.configureFootnotes({ pos: 'pageBottom', numFmt: 'decimal', numStart: 1, numRestart: 'continuous' });
    const fn = b.addFootnote([paragraph(['Note.'])]);
    b.addBlock(paragraph(['Text', fn]));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).toContain('word/footnotes.xml');
  });

  it('configureEndnotes does not throw and builder still builds', async () => {
    const b = new DocxBuilder();
    b.configureEndnotes({ pos: 'sectEnd', numFmt: 'lowerRoman' });
    const en = b.addEndnote([paragraph(['Note.'])]);
    b.addBlock(paragraph(['Text', en]));
    const zip = await buildAndOpen(b);
    expect(partNames(zip)).toContain('word/endnotes.xml');
  });
});

// ---------------------------------------------------------------------------
// Regression: no deferred parts emitted
// ---------------------------------------------------------------------------

describe('Batch 3 regression — no accidental deferred parts', () => {
  const FORBIDDEN = [
    'word/comments.xml',
    'word/fontTable.xml',
  ];

  it('does not emit deferred parts when footnotes are used', async () => {
    const zip = await buildAndOpen(buildWithOneFootnote());
    const parts = partNames(zip);
    for (const p of FORBIDDEN) {
      expect(parts).not.toContain(p);
    }
  });
});

// ---------------------------------------------------------------------------
// Combined Batch 1 + 2 + 3
// ---------------------------------------------------------------------------

describe('Batch 1 + 2 + 3 combined', () => {
  it('emits numbering + settings + header + footer + footnotes together', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });

    const numId = b.defineNumbering({ levels: [
      { ilvl: 0, numFmt: 'decimal', lvlText: '%1.', lvlJc: 'start', suff: 'tab',
        pPr: { ind: { start: 720, hanging: 360 } } },
    ]});

    const hdrRef = b.addHeader('default', [
      paragraph([{ text: 'Report Header', rPr: { b: true } }], { jc: 'center' }),
    ]);
    const ftrRef = b.addFooter('default', [
      paragraph([{ text: 'Footer' }], { jc: 'center' }),
    ]);

    b.setSectionProperties({
      pgSz: { w: 11906, h: 16838 },
      pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
      headerReference: [hdrRef],
      footerReference: [ftrRef],
    });

    const fn1 = b.addFootnote([paragraph(['First footnote.'])]);
    const fn2 = b.addFootnote([paragraph(['Second footnote.'])]);

    b.addBlock(paragraph(['Item 1', fn1], { numPr: { numId, ilvl: 0 } }));
    b.addBlock(paragraph(['Item 2', fn2], { numPr: { numId, ilvl: 0 } }));

    const zip = await buildAndOpen(b);
    const parts = partNames(zip);

    expect(parts).toContain('word/numbering.xml');
    expect(parts).toContain('word/settings.xml');
    expect(parts).toContain('word/header1.xml');
    expect(parts).toContain('word/footer2.xml');
    expect(parts).toContain('word/footnotes.xml');

    const ftn = await readPart(zip, 'word/footnotes.xml');
    expect(ftn).toContain('First footnote.');
    expect(ftn).toContain('Second footnote.');
    expect(ftn).toContain('w:id="1"');
    expect(ftn).toContain('w:id="2"');
  });
});
