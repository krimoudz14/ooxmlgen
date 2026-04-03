/**
 * Batch 1 example — numbering, page numbering, settings, section mechanics,
 *                    table cell completion.
 *
 * Demonstrates reusable mechanisms only. No thesis-specific content.
 * This example could be the foundation for a report, thesis, or any
 * structured document.
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder, paragraph, emptyParagraph, pageBreak,
  ptToTwips, ptToHalfPt, inToTwips,
  bulletListLevel, decimalListLevel, arabicAbjadListLevel,
} from '../src/index';

async function main() {
  const builder = new DocxBuilder();

  builder.setCoreProperties({
    title: 'Batch 1 Mechanisms Example',
    creator: 'ooxmlgen',
    created: new Date(),
    modified: new Date(),
  });

  // Settings — updateFields required for any document with fields
  builder.configureSettings({
    updateFields: true,
    defaultTabStop: 720,
  });

  // A4 page
  builder.setPageSize({ w: 11906, h: 16838 });
  builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });

  // Styles
  builder.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: { spacing: { after: ptToTwips(8), line: 276, lineRule: 'auto' } },
    rPr: { sz: ptToHalfPt(11), rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' } },
  });
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: { outlineLvl: 0, keepNext: true, spacing: { before: ptToTwips(12), after: ptToTwips(6) } },
    rPr: { b: true, sz: ptToHalfPt(14), color: { val: '2E74B5' } },
  });

  // ---- Section 1: Bullet list ----
  builder.addBlock(paragraph([{ text: '1. Bullet List', rPr: {} }], { pStyle: 'Heading1' }));

  const bulletNumId = builder.defineNumbering({
    multiLevelType: 'singleLevel',
    levels: [bulletListLevel(0)],
  });

  builder.addBlock(paragraph(['First bullet item'], { numPr: { numId: bulletNumId, ilvl: 0 } }));
  builder.addBlock(paragraph(['Second bullet item'], { numPr: { numId: bulletNumId, ilvl: 0 } }));
  builder.addBlock(paragraph(['Third bullet item'], { numPr: { numId: bulletNumId, ilvl: 0 } }));

  // ---- Section 2: Decimal numbered list ----
  builder.addBlock(paragraph([{ text: '2. Numbered List', rPr: {} }], { pStyle: 'Heading1' }));

  const decimalNumId = builder.defineNumbering({
    multiLevelType: 'singleLevel',
    levels: [decimalListLevel(0)],
  });

  builder.addBlock(paragraph(['First numbered item'], { numPr: { numId: decimalNumId, ilvl: 0 } }));
  builder.addBlock(paragraph(['Second numbered item'], { numPr: { numId: decimalNumId, ilvl: 0 } }));
  builder.addBlock(paragraph(['Third numbered item'], { numPr: { numId: decimalNumId, ilvl: 0 } }));

  // ---- Section 3: Multi-level list ----
  builder.addBlock(paragraph([{ text: '3. Multi-Level List', rPr: {} }], { pStyle: 'Heading1' }));

  const multiNumId = builder.defineNumbering({
    multiLevelType: 'multilevel',
    levels: [
      decimalListLevel(0),
      { ilvl: 1, numFmt: 'lowerLetter', lvlText: '%2.', lvlJc: 'start', suff: 'tab',
        pPr: { ind: { start: 1440, hanging: 360 } } },
      { ilvl: 2, numFmt: 'lowerRoman', lvlText: '%3.', lvlJc: 'start', suff: 'tab',
        pPr: { ind: { start: 2160, hanging: 360 } } },
    ],
  });

  builder.addBlock(paragraph(['Level 1 item'], { numPr: { numId: multiNumId, ilvl: 0 } }));
  builder.addBlock(paragraph(['Level 2 item'], { numPr: { numId: multiNumId, ilvl: 1 } }));
  builder.addBlock(paragraph(['Level 3 item'], { numPr: { numId: multiNumId, ilvl: 2 } }));
  builder.addBlock(paragraph(['Back to level 1'], { numPr: { numId: multiNumId, ilvl: 0 } }));

  // ---- Section 4: Page numbering — section break with Roman numerals ----
  builder.addBlock(paragraph([{ text: '4. Section Break with Page Numbering', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph(['This paragraph ends section 1 (Roman numerals i, ii, iii).']));

  // Mid-document section break: next section starts with Arabic page 1
  builder.addBlock(paragraph([''], {
    sectPr: {
      type: 'nextPage',
      pgNumType: { fmt: 'upperRoman', start: 1 },
      pgSz: { w: 11906, h: 16838 },
      pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
    },
  }));

  builder.addBlock(paragraph([{ text: 'Section 2 — Arabic Page Numbering', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph(['This section uses Arabic (decimal) page numbering starting at 1.']));

  // ---- Section 5: Table with cell borders and vMerge ----
  builder.addBlock(paragraph([{ text: '5. Table with Cell Borders and Vertical Merge', rPr: {} }], { pStyle: 'Heading1' }));

  const border = { val: 'single', sz: 4, space: 0, color: '4472C4' };
  builder.addTable({
    tblPr: {
      tblW: { w: 8000, type: 'dxa' },
      tblBorders: { top: border, bottom: border, start: border, end: border, insideH: border, insideV: border },
    },
    colWidths: [2000, 3000, 3000],
    rows: [
      // Header row
      {
        cells: [
          { tcPr: { shd: { val: 'clear', fill: 'D9E1F2' } }, content: [paragraph([{ text: 'ID', rPr: { b: true } }])] },
          { tcPr: { shd: { val: 'clear', fill: 'D9E1F2' } }, content: [paragraph([{ text: 'Name', rPr: { b: true } }])] },
          { tcPr: { shd: { val: 'clear', fill: 'D9E1F2' } }, content: [paragraph([{ text: 'Value', rPr: { b: true } }])] },
        ],
      },
      // Row with vMerge restart
      {
        cells: [
          { tcPr: { vMerge: 'restart', tcBorders: { top: border, start: border, bottom: border, end: border } },
            content: [paragraph(['Merged'])] },
          { content: [paragraph(['Alpha'])] },
          { content: [paragraph(['100'])] },
        ],
      },
      // Row with vMerge continue
      {
        cells: [
          { tcPr: { vMerge: 'continue' }, content: [emptyParagraph()] },
          { content: [paragraph(['Beta'])] },
          { content: [paragraph(['200'])] },
        ],
      },
      // Row with cell margins
      {
        cells: [
          { tcPr: { tcMar: { top: { w: 100, type: 'dxa' }, start: { w: 200, type: 'dxa' } } },
            content: [paragraph(['3'])] },
          { content: [paragraph(['Gamma'])] },
          { content: [paragraph(['300'])] },
        ],
      },
    ],
  });

  builder.addBlock(emptyParagraph());

  // ---- Section 6: Arabic list (arabicAbjad) ----
  builder.addBlock(paragraph([{ text: '6. Arabic Abjad List', rPr: {} }], { pStyle: 'Heading1' }));

  const arabicNumId = builder.defineNumbering({
    multiLevelType: 'singleLevel',
    levels: [arabicAbjadListLevel(0)],
  });

  builder.addBlock(paragraph(['أول عنصر'], { numPr: { numId: arabicNumId, ilvl: 0 }, bidi: true, jc: 'start' }));
  builder.addBlock(paragraph(['ثاني عنصر'], { numPr: { numId: arabicNumId, ilvl: 0 }, bidi: true, jc: 'start' }));
  builder.addBlock(paragraph(['ثالث عنصر'], { numPr: { numId: arabicNumId, ilvl: 0 }, bidi: true, jc: 'start' }));

  // Final section properties (Arabic page numbering for body)
  builder.setPageSize({ w: 11906, h: 16838 });
  builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });

  const buf = await builder.build();
  const out = path.join(__dirname, 'output-batch1.docx');
  fs.writeFileSync(out, buf);
  console.log(`Written: ${out} (${buf.length} bytes)`);
}

main().catch(err => { console.error(err); process.exit(1); });
