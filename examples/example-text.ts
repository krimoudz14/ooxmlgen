/**
 * Example 1: Simple text-only document.
 * Exercises: plain paragraphs, mixed run formatting, alignment,
 *            spacing, indentation, page break, section properties.
 * No tables, no styles beyond defaults.
 */

import * as fs from 'fs';
import * as path from 'path';
import { DocxBuilder, paragraph, pageBreak, ptToTwips, ptToHalfPt, inToTwips } from '../src/index';

async function main() {
  const builder = new DocxBuilder();

  builder.setCoreProperties({ title: 'Text-Only Example', creator: 'ooxmlgen' });

  // A4, standard margins
  builder.setPageSize({ w: 11906, h: 16838 });
  builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });

  // Plain paragraph
  builder.addBlock(paragraph(['Hello, World. This is a plain paragraph with no special formatting.']));

  // Mixed inline formatting — no empty rPr should appear
  builder.addBlock(paragraph([
    'Normal text, ',
    { text: 'bold', rPr: { b: true } },
    ', ',
    { text: 'italic', rPr: { i: true } },
    ', ',
    { text: 'bold-italic', rPr: { b: true, i: true } },
    ', ',
    { text: 'underline', rPr: { u: 'single' } },
    ', ',
    { text: 'strikethrough', rPr: { strike: true } },
    ', ',
    { text: 'superscript', rPr: { vertAlign: 'superscript' } },
    ', ',
    { text: 'subscript', rPr: { vertAlign: 'subscript' } },
    '.',
  ]));

  // Font size variation
  builder.addBlock(paragraph([
    { text: '8pt ', rPr: { sz: ptToHalfPt(8) } },
    { text: '10pt ', rPr: { sz: ptToHalfPt(10) } },
    { text: '12pt ', rPr: { sz: ptToHalfPt(12) } },
    { text: '16pt ', rPr: { sz: ptToHalfPt(16) } },
    { text: '24pt', rPr: { sz: ptToHalfPt(24) } },
  ]));

  // Color
  builder.addBlock(paragraph([
    { text: 'Red ', rPr: { color: { val: 'FF0000' } } },
    { text: 'Green ', rPr: { color: { val: '00AA00' } } },
    { text: 'Blue', rPr: { color: { val: '0000FF' } } },
  ]));

  // Alignment
  builder.addBlock(paragraph(['Left aligned (default).'], { jc: 'start' }));
  builder.addBlock(paragraph(['Centered paragraph.'], { jc: 'center' }));
  builder.addBlock(paragraph(['Right aligned paragraph.'], { jc: 'end' }));
  builder.addBlock(paragraph(['Justified paragraph. This text is long enough to demonstrate justification across the full line width of the page.'], { jc: 'both' }));

  // Spacing
  builder.addBlock(paragraph(['Paragraph with 24pt before and 12pt after spacing.'], {
    spacing: { before: ptToTwips(24), after: ptToTwips(12) },
  }));

  // Indentation
  builder.addBlock(paragraph(['First-line indent of 0.5 inches.'], { ind: { firstLine: inToTwips(0.5) } }));
  builder.addBlock(paragraph(['Hanging indent of 0.5 inches.'], { ind: { hanging: inToTwips(0.5), start: inToTwips(0.5) } }));
  builder.addBlock(paragraph(['Left indent of 1 inch.'], { ind: { start: inToTwips(1) } }));

  // Page break
  builder.addBlock(pageBreak());

  // Page 2
  builder.addBlock(paragraph(['This is the first paragraph on page 2, after a page break.']));
  builder.addBlock(paragraph(['Language tag example.'], {
    // pPr only — no rPr needed
  }));
  builder.addBlock(paragraph([
    { text: 'Text with lang=en-US', rPr: { lang: { val: 'en-US' } } },
  ]));

  const buf = await builder.build();
  const out = path.join(__dirname, 'output-text.docx');
  fs.writeFileSync(out, buf);
  console.log(`Written: ${out} (${buf.length} bytes)`);
}

main().catch(err => { console.error(err); process.exit(1); });
