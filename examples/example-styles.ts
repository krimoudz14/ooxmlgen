/**
 * Example 2: Styles-focused document.
 * Exercises: paragraph styles, character styles, style inheritance
 *            (basedOn, next), default style, multiple heading levels.
 */

import * as fs from 'fs';
import * as path from 'path';
import { DocxBuilder, paragraph, emptyParagraph, ptToTwips, ptToHalfPt } from '../src/index';

async function main() {
  const builder = new DocxBuilder();

  builder.setCoreProperties({ title: 'Styles Example', creator: 'ooxmlgen' });
  builder.setPageSize({ w: 11906, h: 16838 });
  builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });

  // ---- Style definitions ----

  // Default paragraph style
  builder.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: { spacing: { after: ptToTwips(8), line: 276, lineRule: 'auto' } },
    rPr: { sz: ptToHalfPt(11), rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' } },
  });

  // Heading 1
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      outlineLvl: 0,
      keepNext: true,
      spacing: { before: ptToTwips(12), after: ptToTwips(3) },
    },
    rPr: { b: true, sz: ptToHalfPt(16), color: { val: '2E74B5' } },
  });

  // Heading 2
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading2', name: 'heading 2',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      outlineLvl: 1,
      keepNext: true,
      spacing: { before: ptToTwips(8), after: ptToTwips(3) },
    },
    rPr: { b: true, sz: ptToHalfPt(13), color: { val: '2E74B5' } },
  });

  // Heading 3
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading3', name: 'heading 3',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      outlineLvl: 2,
      keepNext: true,
      spacing: { before: ptToTwips(6), after: ptToTwips(3) },
    },
    rPr: { b: true, i: true, sz: ptToHalfPt(12), color: { val: '1F3864' } },
  });

  // Block quote style
  builder.addStyle({
    type: 'paragraph', styleId: 'BlockQuote', name: 'Block Quote',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      ind: { start: ptToTwips(36), end: ptToTwips(36) },
      spacing: { before: ptToTwips(6), after: ptToTwips(6) },
    },
    rPr: { i: true, color: { val: '595959' } },
  });

  // Code style
  builder.addStyle({
    type: 'paragraph', styleId: 'Code', name: 'Code',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      spacing: { before: ptToTwips(6), after: ptToTwips(6) },
      shd: { val: 'clear', fill: 'F2F2F2' },
    },
    rPr: { sz: ptToHalfPt(10), rFonts: { ascii: 'Courier New', hAnsi: 'Courier New' } },
  });

  // Character style: Strong
  builder.addStyle({
    type: 'character', styleId: 'Strong', name: 'Strong',
    rPr: { b: true },
  });

  // Character style: Emphasis
  builder.addStyle({
    type: 'character', styleId: 'Emphasis', name: 'Emphasis',
    rPr: { i: true },
  });

  // ---- Content using styles ----

  builder.addBlock(paragraph([{ text: 'Chapter 1: Introduction', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph(['This document demonstrates paragraph and character styles in ooxmlgen v1.']));

  builder.addBlock(paragraph([{ text: '1.1 Background', rPr: {} }], { pStyle: 'Heading2' }));
  builder.addBlock(paragraph(['The following is a block quote demonstrating the BlockQuote style:']));
  builder.addBlock(paragraph(['This is a block quote. It uses italic text and indentation to visually distinguish it from body text.'], { pStyle: 'BlockQuote' }));

  builder.addBlock(paragraph([{ text: '1.2 Code Example', rPr: {} }], { pStyle: 'Heading2' }));
  builder.addBlock(paragraph(['Below is a code sample:']));
  builder.addBlock(paragraph(['const doc = new DocxBuilder();'], { pStyle: 'Code' }));
  builder.addBlock(paragraph(['doc.addBlock(paragraph(["Hello, World!"]));'], { pStyle: 'Code' }));
  builder.addBlock(paragraph(['const buf = await doc.build();'], { pStyle: 'Code' }));

  builder.addBlock(paragraph([{ text: 'Chapter 2: Heading Hierarchy', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph([{ text: '2.1 Section', rPr: {} }], { pStyle: 'Heading2' }));
  builder.addBlock(paragraph([{ text: '2.1.1 Subsection', rPr: {} }], { pStyle: 'Heading3' }));
  builder.addBlock(paragraph(['Body text under a third-level heading. The heading hierarchy is defined by outlineLvl values 0, 1, 2 in the style definitions.']));

  builder.addBlock(paragraph([{ text: '2.2 Style Inheritance', rPr: {} }], { pStyle: 'Heading2' }));
  builder.addBlock(paragraph([
    'All heading styles use ',
    { text: 'basedOn: Normal', rPr: { rFonts: { ascii: 'Courier New', hAnsi: 'Courier New' }, sz: ptToHalfPt(10) } },
    ' and ',
    { text: 'next: Normal', rPr: { rFonts: { ascii: 'Courier New', hAnsi: 'Courier New' }, sz: ptToHalfPt(10) } },
    ' to inherit base formatting and return to Normal after the heading.',
  ]));

  builder.addBlock(emptyParagraph());

  const buf = await builder.build();
  const out = path.join(__dirname, 'output-styles.docx');
  fs.writeFileSync(out, buf);
  console.log(`Written: ${out} (${buf.length} bytes)`);
}

main().catch(err => { console.error(err); process.exit(1); });
