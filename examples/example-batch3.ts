/**
 * Batch 3 example — Footnotes and Endnotes
 *
 * Demonstrates:
 * - addFootnote(): insert footnote references in body text
 * - addEndnote(): insert endnote references in body text
 * - Multi-paragraph note content
 * - configureFootnotes(): position and number format
 * - Combined use with headers, footers, and numbering (Batch 1+2+3)
 *
 * Output: output/batch3-footnotes.docx
 *
 * Inspect with:
 *   unzip -l output/batch3-footnotes.docx
 *   unzip -p output/batch3-footnotes.docx word/footnotes.xml
 *   unzip -p output/batch3-footnotes.docx word/endnotes.xml
 *   unzip -p output/batch3-footnotes.docx word/_rels/document.xml.rels
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder,
  paragraph,
  textRun,
  pageNumberField,
  ptToTwips,
  inToTwips,
} from '../src/index';

async function main() {
  const b = new DocxBuilder();

  // ---- Document properties ----
  b.setCoreProperties({
    title: 'Batch 3 — Footnotes and Endnotes Demo',
    creator: 'ooxmlgen',
  });

  // ---- Page setup ----
  b.setPageSize({ w: 11906, h: 16838 }); // A4
  b.setPageMargins({
    top: inToTwips(1), right: inToTwips(1.25),
    bottom: inToTwips(1), left: inToTwips(1.25),
    header: 720, footer: 720, gutter: 0,
  });

  // ---- Settings ----
  b.configureSettings({ updateFields: true });

  // ---- Footer with page number ----
  const ftrRef = b.addFooter('default', [
    paragraph([...pageNumberField()], { jc: 'center' }),
  ]);
  b.setSectionProperties({
    pgSz: { w: 11906, h: 16838 },
    pgMar: {
      top: inToTwips(1), right: inToTwips(1.25),
      bottom: inToTwips(1), left: inToTwips(1.25),
      header: 720, footer: 720, gutter: 0,
    },
    footerReference: [ftrRef],
  });

  // ---- Configure footnote behavior ----
  b.configureFootnotes({
    pos: 'pageBottom',
    numFmt: 'decimal',
    numStart: 1,
    numRestart: 'continuous',
  });

  // ---- Body content with footnotes ----

  // Title
  b.addBlock(paragraph(
    [{ text: 'Footnotes and Endnotes Demo', rPr: { b: true, sz: ptToTwips(1) * 14 } }],
    { jc: 'center', spacing: { after: 400 } },
  ));

  // Paragraph 1 — single footnote
  const fn1 = b.addFootnote([
    paragraph(['This is the first footnote. It appears at the bottom of the page.']),
  ]);
  b.addBlock(paragraph([
    'The first paragraph contains a footnote reference',
    fn1,
    ' at the end of this sentence.',
  ], { spacing: { after: 200 } }));

  // Paragraph 2 — footnote with multi-run content
  const fn2 = b.addFootnote([
    paragraph([
      { text: 'Second footnote: ', rPr: { b: true } },
      'This note has bold and plain text mixed together.',
    ]),
  ]);
  b.addBlock(paragraph([
    'The second paragraph has another reference',
    fn2,
    ' mid-sentence, demonstrating sequential IDs.',
  ], { spacing: { after: 200 } }));

  // Paragraph 3 — two footnotes in one paragraph
  const fn3 = b.addFootnote([
    paragraph(['Third footnote — first reference in this paragraph.']),
  ]);
  const fn4 = b.addFootnote([
    paragraph(['Fourth footnote — second reference in this paragraph.']),
  ]);
  b.addBlock(paragraph([
    'This paragraph has two references',
    fn3,
    ' and another one',
    fn4,
    ' in the same paragraph.',
  ], { spacing: { after: 400 } }));

  // Section heading
  b.addBlock(paragraph(
    [{ text: 'Endnotes Section', rPr: { b: true, sz: ptToTwips(1) * 13 } }],
    { spacing: { before: 200, after: 200 } },
  ));

  // Paragraph with endnote
  const en1 = b.addEndnote([
    paragraph(['This is an endnote. Endnotes appear at the end of the document or section.']),
  ]);
  b.addBlock(paragraph([
    'This paragraph uses an endnote reference',
    en1,
    ' instead of a footnote.',
  ], { spacing: { after: 200 } }));

  // Paragraph with both footnote and endnote
  const fn5 = b.addFootnote([
    paragraph(['Fifth footnote — mixed with an endnote in the same paragraph.']),
  ]);
  const en2 = b.addEndnote([
    paragraph(['Second endnote — IDs are tracked independently from footnotes.']),
  ]);
  b.addBlock(paragraph([
    'This paragraph mixes a footnote',
    fn5,
    ' and an endnote',
    en2,
    ' to show independent ID sequences.',
  ], { spacing: { after: 200 } }));

  // ---- Build and write ----
  const outDir = path.join(__dirname, '..', 'output');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outPath = path.join(outDir, 'batch3-footnotes.docx');
  const buf = await b.build();
  fs.writeFileSync(outPath, buf);

  console.log(`Written: ${outPath}`);
  console.log(`Size: ${buf.length} bytes`);
  console.log('');
  console.log('Inspect with:');
  console.log(`  unzip -l "${outPath}"`);
  console.log(`  unzip -p "${outPath}" word/footnotes.xml`);
  console.log(`  unzip -p "${outPath}" word/endnotes.xml`);
  console.log(`  unzip -p "${outPath}" word/_rels/document.xml.rels`);
}

main().catch(err => { console.error(err); process.exit(1); });
