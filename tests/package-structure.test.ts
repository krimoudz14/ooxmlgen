/**
 * Package structure tests — fixture-based validation of generated DOCX files.
 *
 * Verifies:
 * 1. Expected v1 parts are present
 * 2. Forbidden deferred parts are absent
 * 3. Content types are correct for each part
 * 4. Relationships are correctly wired
 */

import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import { DocxBuilder, paragraph, emptyParagraph } from '../src/index';

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const EXPECTED_PARTS = [
  '[Content_Types].xml',
  '_rels/.rels',
  'word/document.xml',
  'word/styles.xml',
  'word/_rels/document.xml.rels',
  'docProps/core.xml',
  'docProps/app.xml',
] as const;

const FORBIDDEN_PARTS = [
  'word/numbering.xml',
  'word/settings.xml',
  'word/fontTable.xml',
  'word/header1.xml',
  'word/header2.xml',
  'word/footer1.xml',
  'word/footer2.xml',
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/comments.xml',
  'word/theme/theme1.xml',
] as const;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function buildMinimal(): Promise<Buffer> {
  const b = new DocxBuilder();
  b.addBlock(paragraph(['Hello']));
  return b.build();
}

async function openZip(buf: Buffer): Promise<JSZip> {
  return JSZip.loadAsync(buf);
}

async function readEntry(zip: JSZip, name: string): Promise<string> {
  const entry = zip.file(name);
  if (!entry) throw new Error(`Entry not found: ${name}`);
  return entry.async('string');
}

function getPartNames(zip: JSZip): string[] {
  return Object.keys(zip.files).filter(n => !zip.files[n].dir);
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('DOCX package structure', () => {
  let buf: Buffer;
  let zip: JSZip;
  let parts: string[];

  beforeAll(async () => {
    buf = await buildMinimal();
    zip = await openZip(buf);
    parts = getPartNames(zip);
  });

  it('produces a non-empty buffer', () => {
    expect(buf.length).toBeGreaterThan(0);
  });

  it('is a valid ZIP (loadAsync does not throw)', () => {
    expect(zip).toBeDefined();
  });

  describe('expected parts', () => {
    for (const part of EXPECTED_PARTS) {
      it(`contains ${part}`, () => {
        expect(parts).toContain(part);
      });
    }
  });

  describe('forbidden deferred parts', () => {
    for (const part of FORBIDDEN_PARTS) {
      it(`does NOT contain ${part}`, () => {
        expect(parts).not.toContain(part);
      });
    }
  });

  it('contains exactly the expected parts (no extras)', () => {
    const unexpected = parts.filter(p => !(EXPECTED_PARTS as readonly string[]).includes(p));
    expect(unexpected).toEqual([]);
  });

  describe('[Content_Types].xml', () => {
    let ct: string;
    beforeAll(async () => { ct = await readEntry(zip, '[Content_Types].xml'); });

    it('has correct namespace', () => {
      expect(ct).toContain('http://schemas.openxmlformats.org/package/2006/content-types');
    });
    it('has Default for rels', () => {
      expect(ct).toContain('Extension="rels"');
      expect(ct).toContain('application/vnd.openxmlformats-package.relationships+xml');
    });
    it('has Default for xml', () => {
      expect(ct).toContain('Extension="xml"');
    });
    it('has Override for word/document.xml', () => {
      expect(ct).toContain('PartName="/word/document.xml"');
      expect(ct).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
    });
    it('has Override for word/styles.xml', () => {
      expect(ct).toContain('PartName="/word/styles.xml"');
      expect(ct).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
    });
    it('has Override for docProps/core.xml', () => {
      expect(ct).toContain('PartName="/docProps/core.xml"');
      expect(ct).toContain('application/vnd.openxmlformats-package.core-properties+xml');
    });
    it('has Override for docProps/app.xml', () => {
      expect(ct).toContain('PartName="/docProps/app.xml"');
      expect(ct).toContain('application/vnd.openxmlformats-officedocument.extended-properties+xml');
    });
  });

  describe('_rels/.rels', () => {
    let rels: string;
    beforeAll(async () => { rels = await readEntry(zip, '_rels/.rels'); });

    it('has correct namespace', () => {
      expect(rels).toContain('http://schemas.openxmlformats.org/package/2006/relationships');
    });
    it('has officeDocument relationship to word/document.xml', () => {
      expect(rels).toContain('relationships/officeDocument');
      expect(rels).toContain('Target="word/document.xml"');
    });
    it('has core-properties relationship', () => {
      expect(rels).toContain('metadata/core-properties');
      expect(rels).toContain('Target="docProps/core.xml"');
    });
    it('has extended-properties relationship', () => {
      expect(rels).toContain('relationships/extended-properties');
      expect(rels).toContain('Target="docProps/app.xml"');
    });
    it('all relationship Ids are unique', () => {
      const ids = [...rels.matchAll(/Id="([^"]+)"/g)].map(m => m[1]);
      expect(new Set(ids).size).toBe(ids.length);
    });
  });

  describe('word/_rels/document.xml.rels', () => {
    let rels: string;
    beforeAll(async () => { rels = await readEntry(zip, 'word/_rels/document.xml.rels'); });

    it('has styles relationship', () => {
      expect(rels).toContain('relationships/styles');
      expect(rels).toContain('Target="styles.xml"');
    });
  });
});

describe('DOCX package structure — with core properties', () => {
  it('emits title and creator in docProps/core.xml', async () => {
    const b = new DocxBuilder();
    b.setCoreProperties({ title: 'Test Title', creator: 'Test Author' });
    b.addBlock(paragraph(['x']));
    const buf = await b.build();
    const zip = await openZip(buf);
    const core = await readEntry(zip, 'docProps/core.xml');
    expect(core).toContain('Test Title');
    expect(core).toContain('Test Author');
  });

  it('emits Application in docProps/app.xml', async () => {
    const b = new DocxBuilder();
    b.setExtendedProperties({ application: 'MyApp', appVersion: '2.0' });
    b.addBlock(paragraph(['x']));
    const buf = await b.build();
    const zip = await openZip(buf);
    const app = await readEntry(zip, 'docProps/app.xml');
    expect(app).toContain('<Application>MyApp</Application>');
    expect(app).toContain('<AppVersion>2.0</AppVersion>');
  });
});
