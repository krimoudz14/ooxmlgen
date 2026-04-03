/**
 * Batch 4 tests — Hyperlinks, TOC, SEQ/caption fields
 *
 * Validates:
 * - External hyperlinks with correct r:id and TargetMode="External"
 * - Internal bookmark anchor hyperlinks
 * - TOC field markup (fldChar + instrText)
 * - SEQ field markup for figure/table numbering
 * - Caption helper produces correct structure
 * - Hyperlink relationships in document.xml.rels
 * - No Batch 5 features emitted accidentally
 *
 * Source: wml.xsd CT_Hyperlink (line 1218), CT_FldChar (line 1210),
 *         TOC/SEQ field instructions (ECMA-376 Part1 §17.16.5)
 */

import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  DocxBuilder,
  paragraph,
  textRun,
  createHyperlink,
  tocField,
  seqField,
  createCaptionParagraph,
  createTableOfContents,
  styleRefField,
  refField,
} from '../src/index';
import type { WmlHyperlink, WmlRun } from '../src/index';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function buildDocx(builder: DocxBuilder): Promise<JSZip> {
  const buf = await builder.build();
  return JSZip.loadAsync(buf);
}

async function getXml(zip: JSZip, path: string): Promise<string | null> {
  const file = zip.file(path);
  if (!file) return null;
  return file.async('string');
}

async function getParts(zip: JSZip): Promise<string[]> {
  return Object.keys(zip.files).filter(f => !zip.files[f].dir);
}

// ---------------------------------------------------------------------------
// Hyperlink tests
// ---------------------------------------------------------------------------

describe('hyperlinks — external URLs', () => {
  it('emits w:hyperlink with r:id in document.xml', async () => {
    const b = new DocxBuilder();
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink',
      rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Visit Example', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('<w:hyperlink');
    expect(docXml).toContain('r:id="');
    expect(docXml).toContain('Visit Example');
  });

  it('emits hyperlink relationship with TargetMode="External"', async () => {
    const b = new DocxBuilder();
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink',
      rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Link', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const rels = await getXml(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain('https://example.com');
    expect(rels).toContain('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
  });

  it('resolves placeholder rId to real rId in document.xml', async () => {
    const b = new DocxBuilder();
    const placeholder = b.addHyperlinkRelationship('https://example.com');
    // placeholder should NOT appear in the final document
    expect(placeholder).toMatch(/^__hlink_\d+__$/);

    const link: WmlHyperlink = {
      _type: 'hyperlink',
      rId: placeholder,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Link', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    // Should contain a real rId (rIdN), not the placeholder
    expect(docXml).not.toContain('__hlink_');
    expect(docXml).toMatch(/r:id="rId\d+"/);
  });

  it('supports multiple hyperlinks with distinct rIds', async () => {
    const b = new DocxBuilder();
    const rId1 = b.addHyperlinkRelationship('https://example.com');
    const rId2 = b.addHyperlinkRelationship('https://other.com');
    expect(rId1).not.toBe(rId2);

    const link1: WmlHyperlink = {
      _type: 'hyperlink', rId: rId1,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'First', space: 'preserve' }] }],
    };
    const link2: WmlHyperlink = {
      _type: 'hyperlink', rId: rId2,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Second', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link1, link2]));
    const zip = await buildDocx(b);
    const rels = await getXml(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('https://example.com');
    expect(rels).toContain('https://other.com');
    // Two distinct hyperlink relationships created
    expect(rId1).not.toBe(rId2);
  });

  it('supports hyperlink with tooltip', async () => {
    const b = new DocxBuilder();
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink', rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Link', space: 'preserve' }] }],
      tooltip: 'Click to visit Example',
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('w:tooltip="Click to visit Example"');
  });

  it('createHyperlink factory produces correct structure', () => {
    const link = createHyperlink({ url: 'https://example.com', text: 'Click here' });
    expect(link._type).toBe('hyperlink');
    expect(link.runs).toHaveLength(1);
    expect(link.runs[0].content[0]).toEqual({ _type: 'text', text: 'Click here', space: 'preserve' });
    // Default hyperlink styling
    expect(link.runs[0].rPr).toEqual({ u: 'single', color: { val: '0563C1' } });
  });

  it('createHyperlink with custom formatting overrides defaults', () => {
    const link = createHyperlink({
      url: 'https://example.com',
      text: 'Custom link',
      rPr: { b: true, u: 'double' },
    });
    expect(link.runs[0].rPr).toEqual({ b: true, u: 'double', color: { val: '0563C1' } });
  });
});

describe('hyperlinks — internal anchors', () => {
  it('emits w:hyperlink with w:anchor for internal bookmark', async () => {
    const b = new DocxBuilder();
    const link: WmlHyperlink = {
      _type: 'hyperlink',
      anchor: 'MyBookmark',
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Go to section', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('w:anchor="MyBookmark"');
    expect(docXml).toContain('Go to section');
    // No r:id for internal anchor
    const hlinkMatch = docXml!.match(/<w:hyperlink[^>]*>/);
    expect(hlinkMatch![0]).not.toContain('r:id');
  });

  it('internal anchor does NOT create a relationship', async () => {
    const b = new DocxBuilder();
    const link: WmlHyperlink = {
      _type: 'hyperlink',
      anchor: 'MyBookmark',
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Go to section', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const rels = await getXml(zip, 'word/_rels/document.xml.rels');
    expect(rels).not.toContain('relationships/hyperlink');
  });
});

// ---------------------------------------------------------------------------
// TOC field tests
// ---------------------------------------------------------------------------

describe('TOC field', () => {
  it('tocField produces complete field sequence', () => {
    const runs = tocField({ headingLevels: [1, 2, 3] });
    // Should have 5 runs: begin, instrText, separate, display, end
    expect(runs).toHaveLength(5);
    // First run: fldChar begin
    expect(runs[0].content[0]).toEqual({ _type: 'fldChar', fldCharType: 'begin' });
    // Second run: instrText
    expect(runs[1].content[0]).toMatchObject({ _type: 'instrText' });
    expect((runs[1].content[0] as any).text).toContain('TOC');
    expect((runs[1].content[0] as any).text).toContain('\\o "1-3"');
    // Third run: fldChar separate
    expect(runs[2].content[0]).toEqual({ _type: 'fldChar', fldCharType: 'separate' });
    // Fourth run: display text
    expect(runs[3].content[0]).toMatchObject({ _type: 'text' });
    // Fifth run: fldChar end
    expect(runs[4].content[0]).toEqual({ _type: 'fldChar', fldCharType: 'end' });
  });

  it('tocField with hyperlinks and page numbers', () => {
    const runs = tocField({ headingLevels: [1, 2, 3], withHyperlinks: true, withPageNumbers: true });
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('\\h');
    expect(instr).toContain('\\z');
  });

  it('tocField with custom switches', () => {
    const runs = tocField({ customSwitches: '\\o "1-3" \\h \\z \\u' });
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('TOC \\o "1-3" \\h \\z \\u');
  });

  it('TOC field serializes correctly in document.xml', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });
    const toc = b.createTableOfContents({ headingLevels: [1, 2, 3] });
    b.addBlock(toc);
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('w:fldCharType="begin"');
    expect(docXml).toContain('w:instrText');
    expect(docXml).toContain('TOC');
    expect(docXml).toContain('w:fldCharType="separate"');
    expect(docXml).toContain('w:fldCharType="end"');
  });
});

// ---------------------------------------------------------------------------
// SEQ field tests
// ---------------------------------------------------------------------------

describe('SEQ field', () => {
  it('seqField produces complete field sequence', () => {
    const runs = seqField('Figure');
    expect(runs).toHaveLength(5);
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('SEQ Figure');
  });

  it('seqField with reset', () => {
    const runs = seqField('Table', { resetTo: 5 });
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('\\r 5');
  });

  it('seqField with nearest preceding', () => {
    const runs = seqField('Equation', { nearestPreceding: true });
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('\\c');
  });

  it('SEQ field serializes correctly in document.xml', async () => {
    const b = new DocxBuilder();
    b.addBlock(paragraph([
      { text: 'Figure ', rPr: { b: true } },
      ...seqField('Figure', { displayText: '1' }),
      { text: ': System architecture', rPr: { i: true } },
    ]));
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('SEQ Figure');
    expect(docXml).toContain('w:fldCharType="begin"');
    expect(docXml).toContain('w:fldCharType="end"');
  });
});

// ---------------------------------------------------------------------------
// Caption helper tests
// ---------------------------------------------------------------------------

describe('caption helper', () => {
  it('createCaptionParagraph produces caption with SEQ field', async () => {
    const b = new DocxBuilder();
    const caption = b.createCaptionParagraph('Figure', 'System architecture diagram');
    b.addBlock(caption);
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('Figure ');
    expect(docXml).toContain('SEQ Figure');
    expect(docXml).toContain(': System architecture diagram');
  });

  it('createCaptionParagraph with custom style', async () => {
    const b = new DocxBuilder();
    b.addStyle({
      type: 'paragraph', styleId: 'Caption', name: 'caption',
      basedOn: 'Normal', next: 'Normal',
    });
    const caption = b.createCaptionParagraph('Table', 'Data summary', { pStyle: 'Caption' });
    b.addBlock(caption);
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('<w:pStyle w:val="Caption"/>');
    expect(docXml).toContain('Table ');
    expect(docXml).toContain('SEQ Table');
    expect(docXml).toContain(': Data summary');
  });

  it('createCaptionParagraph with custom style', async () => {
    const b = new DocxBuilder();
    b.addStyle({
      type: 'paragraph', styleId: 'Caption', name: 'caption',
      basedOn: 'Normal', next: 'Normal',
    });
    const caption = b.createCaptionParagraph('Table', 'Data summary', { pStyle: 'Caption' });
    b.addBlock(caption);
    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');
    expect(docXml).toContain('<w:pStyle w:val="Caption"/>');
    expect(docXml).toContain('Table ');
    expect(docXml).toContain('SEQ Table');
    expect(docXml).toContain(': Data summary');
  });
});

// ---------------------------------------------------------------------------
// Other field helpers
// ---------------------------------------------------------------------------

describe('other field helpers', () => {
  it('styleRefField produces STYLEREF instruction', () => {
    const runs = styleRefField('Heading1');
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('STYLEREF "Heading1"');
  });

  it('refField produces REF instruction', () => {
    const runs = refField('MyBookmark');
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('REF MyBookmark');
  });

  it('refField with paragraph number switch', () => {
    const runs = refField('MyBookmark', { insertParagraphNumber: true });
    const instr = (runs[1].content[0] as any).text;
    expect(instr).toContain('\\p');
  });
});

// ---------------------------------------------------------------------------
// Batch 4 regression — no accidental deferred parts
// ---------------------------------------------------------------------------

describe('Batch 4 regression — no accidental deferred parts', () => {
  it('does not emit Batch 5 features (docDefaults, fontTable) with hyperlinks', async () => {
    const b = new DocxBuilder();
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink', rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Link', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const zip = await buildDocx(b);
    const parts = await getParts(zip);
    expect(parts).not.toContain('word/fontTable.xml');
  });

  it('hyperlinks + footnotes + numbering all work together', async () => {
    const b = new DocxBuilder();

    // Batch 1: numbering
    const numId = b.defineNumbering({
      multiLevelType: 'singleLevel',
      levels: [{ ilvl: 0, numFmt: 'decimal', lvlText: '%1.', lvlJc: 'start',
                  pPr: { ind: { start: 720, hanging: 360 } } }],
    });

    // Batch 3: footnote
    const fn = b.addFootnote([paragraph(['A footnote.'])]);

    // Batch 4: hyperlink
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink', rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'See reference', space: 'preserve' }] }],
    };

    b.addBlock(paragraph(['Item 1', fn, link], { numPr: { numId, ilvl: 0 } }));

    const zip = await buildDocx(b);
    const parts = await getParts(zip);

    expect(parts).toContain('word/footnotes.xml');
    expect(parts).toContain('word/numbering.xml');
    expect(parts).not.toContain('word/fontTable.xml');
    expect(parts).not.toContain('word/comments.xml');

    // Verify hyperlink relationship
    const rels = await getXml(zip, 'word/_rels/document.xml.rels');
    expect(rels).toContain('relationships/hyperlink');
    expect(rels).toContain('https://example.com');
  });

  it('TOC + SEQ + hyperlinks all work together', async () => {
    const b = new DocxBuilder();
    b.configureSettings({ updateFields: true });

    // TOC
    const toc = b.createTableOfContents({ headingLevels: [1, 2, 3] });
    b.addBlock(toc);

    // Caption with SEQ
    const caption = b.createCaptionParagraph('Figure', 'System diagram');
    b.addBlock(caption);

    // Hyperlink
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink', rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'More info', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));

    const zip = await buildDocx(b);
    const docXml = await getXml(zip, 'word/document.xml');

    // TOC field
    expect(docXml).toContain('TOC');
    // SEQ field
    expect(docXml).toContain('SEQ Figure');
    // Hyperlink
    expect(docXml).toContain('<w:hyperlink');
  });
});

// ---------------------------------------------------------------------------
// XML validity
// ---------------------------------------------------------------------------

describe('Batch 4 XML validity', () => {
  it('all parts are well-formed XML with hyperlinks', async () => {
    const b = new DocxBuilder();
    const rId = b.addHyperlinkRelationship('https://example.com');
    const link: WmlHyperlink = {
      _type: 'hyperlink', rId,
      runs: [{ _type: 'run', content: [{ _type: 'text', text: 'Link', space: 'preserve' }] }],
    };
    b.addBlock(paragraph([link]));
    const caption = b.createCaptionParagraph('Table', 'Test table');
    b.addBlock(caption);
    const toc = b.createTableOfContents({ headingLevels: [1, 2] });
    b.addBlock(toc);
    const zip = await buildDocx(b);

    for (const [path, file] of Object.entries(zip.files)) {
      if (file.dir || (!path.endsWith('.xml') && !path.endsWith('.rels'))) continue;
      const content = await file.async('string');
      const opens = (content.match(/</g) || []).length;
      const closes = (content.match(/>/g) || []).length;
      expect(opens).toBe(closes);
    }
  });
});
