/**
 * Batch 2 example — headers, footers, page number fields.
 * Demonstrates reusable mechanisms. No document-type-specific content.
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder, paragraph, emptyParagraph,
  pageNumberField, totalPagesField, pageXofYRuns,
  ptToHalfPt, ptToTwips,
} from '../src/index';

async function main() {
  const builder = new DocxBuilder();

  builder.setCoreProperties({ title: 'Batch 2 — Headers, Footers, Page Numbers', creator: 'ooxmlgen' });
  builder.configureSettings({ updateFields: true, evenAndOddHeaders: true });

  // Styles
  builder.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: { spacing: { after: ptToTwips(8) } },
    rPr: { sz: ptToHalfPt(11), rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' } },
  });
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: { outlineLvl: 0, spacing: { before: ptToTwips(12), after: ptToTwips(6) } },
    rPr: { b: true, sz: ptToHalfPt(14) },
  });

  // Default header: document title centered
  const defaultHdr = builder.addHeader('default', [
    paragraph([{ text: 'Batch 2 Mechanisms Example', rPr: { b: true } }], { jc: 'center' }),
  ]);

  // First-page header: blank (common pattern)
  const firstHdr = builder.addHeader('first', [emptyParagraph()]);

  // Even-page header: right-aligned
  const evenHdr = builder.addHeader('even', [
    paragraph([{ text: 'Even Page Header', rPr: { i: true } }], { jc: 'end' }),
  ]);

  // Default footer: page X of Y centered
  const defaultFtr = builder.addFooter('default', [
    paragraph([...pageXofYRuns()], { jc: 'center' }),
  ]);

  // First-page footer: blank
  const firstFtr = builder.addFooter('first', [emptyParagraph()]);

  // Even-page footer: page number right-aligned
  const evenFtr = builder.addFooter('even', [
    paragraph([...pageNumberField()], { jc: 'end' }),
  ]);

  // Set section properties with all header/footer references
  builder.setSectionProperties({
    pgSz: { w: 11906, h: 16838 },
    pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
    titlePg: true,
    headerReference: [defaultHdr, firstHdr, evenHdr],
    footerReference: [defaultFtr, firstFtr, evenFtr],
  });

  // Content
  builder.addBlock(paragraph([{ text: 'Section 1: Introduction', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph(['This document demonstrates header and footer support in ooxmlgen v2.']));
  builder.addBlock(paragraph(['The default header shows the document title. The footer shows page X of Y.']));
  builder.addBlock(paragraph(['The first page has a blank header and footer (titlePg=true).']));
  builder.addBlock(paragraph(['Even pages have a different header and footer (evenAndOddHeaders=true in settings).']));

  const buf = await builder.build();
  const out = path.join(__dirname, 'output-batch2.docx');
  fs.writeFileSync(out, buf);
  console.log(`Written: ${out} (${buf.length} bytes)`);
}

main().catch(err => { console.error(err); process.exit(1); });
