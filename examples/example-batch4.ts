/**
 * Batch 4 Example — Hyperlinks, TOC, SEQ/caption fields
 *
 * Demonstrates:
 * - External hyperlinks with r:id resolution
 * - Internal anchor hyperlinks
 * - TOC field generation
 * - SEQ field captions
 * - All Batch 1-4 features combined
 */

import * as path from 'path';
import * as fs from 'fs';
import {
  DocxBuilder,
  paragraph,
  textRun,
  pageBreak,
  createHyperlink,
  tocField,
  seqField,
  pageXofYRuns,
  styleRefField,
  inToTwips,
  ptToHalfPt,
} from '../src/index';

async function main() {
  const b = new DocxBuilder();

  // Batch 2: header and footer
  const hdrRef = b.addHeader('default', [
    paragraph([{ text: 'ooxmlgen — Batch 4 Demo', rPr: { i: true } }], { jc: 'center' }),
  ]);
  const ftrRef = b.addFooter('default', [
    paragraph([...pageXofYRuns()], { jc: 'center' }),
  ]);

  // Batch 1: numbering
  const numId = b.defineNumbering({
    multiLevelType: 'singleLevel',
    levels: [{ ilvl: 0, numFmt: 'decimal', lvlText: '%1.', lvlJc: 'start',
                pPr: { ind: { start: 720, hanging: 360 } } }],
  });

  b.configureSettings({ updateFields: true });

  // Title
  b.addBlock(paragraph([{ text: 'Batch 4: Advanced Word Mechanisms' }], {
    pStyle: 'Heading1',
  }));

  // Section 1: Hyperlinks
  b.addBlock(paragraph([{ text: '1. Hyperlinks' }], { pStyle: 'Heading2' }));

  // External hyperlink
  const extRId = b.addHyperlinkRelationship('https://github.com');
  const extLink = createHyperlink({
    url: 'https://github.com',
    text: 'Visit GitHub',
  });
  extLink.rId = extRId; // Set the resolved rId
  b.addBlock(paragraph(['Check out the source code at ', extLink, '.']));

  // Internal anchor hyperlink
  const intLink = createHyperlink({
    anchor: 'toc-section',
    text: 'Jump to TOC',
  });
  b.addBlock(paragraph(['Or ', intLink, ' for the table of contents.']));

  // Section 2: TOC
  b.addBlock(paragraph([{ text: '2. Table of Contents' }], { pStyle: 'Heading2' }));
  const tocPara = b.createTableOfContents({ headingLevels: [1, 2, 3] });
  b.addBlock(tocPara);

  // Section 3: Captions and SEQ fields
  b.addBlock(paragraph([{ text: '3. Figures and Tables' }], { pStyle: 'Heading2' }));

  const figCaption = b.createCaptionParagraph('Figure', 'System architecture diagram');
  b.addBlock(figCaption);

  const tblCaption = b.createCaptionParagraph('Table', 'Data summary table', {
    pStyle: 'Caption',
  });
  b.addBlock(tblCaption);

  // Section 4: Field examples
  b.addBlock(paragraph([{ text: '4. Advanced Fields' }], { pStyle: 'Heading2' }));

  // Style reference (for headers)
  const styleRefRuns = styleRefField('Heading1');
  b.addBlock(paragraph(['This document is titled: ', ...styleRefRuns, '.']));

  // Cross-reference to figure
  const refRuns = seqField('Figure', { displayText: '1' });
  b.addBlock(paragraph(['See ', ...refRuns, ' for the architecture diagram.']));

  // Page break
  b.addBlock(pageBreak());

  // Section 5: Combined features
  b.addBlock(paragraph([{ text: '5. Combined Features' }], { pStyle: 'Heading2' }));

  // Numbered list with footnote and hyperlink
  const fn = b.addFootnote([paragraph(['This footnote contains a link to the repository.'])]);
  b.addBlock(paragraph(['Item with footnote and hyperlink', fn], { numPr: { numId, ilvl: 0 } }));

  // Styles
  b.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: { spacing: { before: 240, after: 120 }, outlineLvl: 0 },
    rPr: { b: true, color: { val: '2E74B5' }, sz: ptToHalfPt(16) },
  });
  b.addStyle({
    type: 'paragraph', styleId: 'Heading2', name: 'heading 2',
    basedOn: 'Normal', next: 'Normal',
    pPr: { spacing: { before: 180, after: 60 }, outlineLvl: 1 },
    rPr: { b: true, sz: ptToHalfPt(14) },
  });
  b.addStyle({
    type: 'paragraph', styleId: 'Caption', name: 'caption',
    basedOn: 'Normal', next: 'Normal',
    pPr: { spacing: { before: 60, after: 60 }, jc: 'center' },
    rPr: { i: true, sz: ptToHalfPt(10) },
  });
  b.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: { spacing: { after: 160, line: 276, lineRule: 'auto' } },
    rPr: { rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' }, sz: ptToHalfPt(11) },
  });

  // Section properties with header and footer
  b.setSectionProperties({
    headerReference: [hdrRef],
    footerReference: [ftrRef],
    pgSz: { w: 11906, h: 16838 },  // A4
    pgMar: {
      top:    inToTwips(1),
      right:  inToTwips(1.25),
      bottom: inToTwips(1),
      left:   inToTwips(1.25),
      header: 720,
      footer: 720,
      gutter: 0,
    },
  });

  b.setCoreProperties({
    title: 'ooxmlgen Batch 4 Demo',
    creator: 'ooxmlgen',
    created: new Date('2026-04-02T00:00:00.000Z'),
    modified: new Date('2026-04-02T00:00:00.000Z'),
  });

  const buf = await b.build();
  const outPath = path.join(__dirname, 'output-batch4.docx');
  fs.writeFileSync(outPath, buf);
  console.log(`Written: ${outPath} (${buf.length} bytes)`);
}

main().catch(console.error);