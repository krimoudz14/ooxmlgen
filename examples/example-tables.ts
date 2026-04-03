/**
 * Example 3: Tables and sections document.
 * Exercises: multiple tables, cell spanning (gridSpan), cell shading,
 *            table width types, section type, landscape orientation,
 *            vertical alignment, row properties.
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder, paragraph, emptyParagraph,
  ptToTwips, ptToHalfPt, inToTwips,
  WmlBlockElement,
} from '../src/index';

// Helper: cell with a single paragraph
function cell(text: string, opts: {
  bold?: boolean; shade?: string; gridSpan?: number; vAlign?: 'top' | 'center' | 'bottom'; w?: number;
} = {}): { tcPr?: any; content: WmlBlockElement[] } {
  return {
    tcPr: {
      ...(opts.gridSpan ? { gridSpan: opts.gridSpan } : {}),
      ...(opts.shade ? { shd: { val: 'clear', fill: opts.shade } } : {}),
      ...(opts.vAlign ? { vAlign: opts.vAlign } : {}),
      ...(opts.w ? { tcW: { w: opts.w, type: 'dxa' as const } } : {}),
    },
    content: [paragraph([{ text, rPr: opts.bold ? { b: true } : undefined }])],
  };
}

async function main() {
  const builder = new DocxBuilder();

  builder.setCoreProperties({ title: 'Tables and Sections Example', creator: 'ooxmlgen' });
  builder.setPageSize({ w: 11906, h: 16838 });
  builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });

  builder.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: { spacing: { after: ptToTwips(6) } },
    rPr: { sz: ptToHalfPt(11), rFonts: { ascii: 'Calibri', hAnsi: 'Calibri' } },
  });
  builder.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: { outlineLvl: 0, spacing: { before: ptToTwips(12), after: ptToTwips(6) } },
    rPr: { b: true, sz: ptToHalfPt(14) },
  });

  // ---- Table 1: Simple bordered table ----
  builder.addBlock(paragraph([{ text: 'Table 1: Simple Bordered Table', rPr: {} }], { pStyle: 'Heading1' }));

  const col3 = Math.round(8406 / 3);
  builder.addTable({
    tblPr: {
      tblW: { w: 8406, type: 'dxa' },
      tblBorders: {
        top:     { val: 'single', sz: 4, space: 0, color: '000000' },
        bottom:  { val: 'single', sz: 4, space: 0, color: '000000' },
        start:   { val: 'single', sz: 4, space: 0, color: '000000' },
        end:     { val: 'single', sz: 4, space: 0, color: '000000' },
        insideH: { val: 'single', sz: 4, space: 0, color: '000000' },
        insideV: { val: 'single', sz: 4, space: 0, color: '000000' },
      },
    },
    colWidths: [col3, col3, col3],
    rows: [
      { cells: [cell('Name', { bold: true, shade: 'D9E1F2' }), cell('Role', { bold: true, shade: 'D9E1F2' }), cell('Status', { bold: true, shade: 'D9E1F2' })] },
      { cells: [cell('Alice'), cell('Engineer'), cell('Active')] },
      { cells: [cell('Bob'), cell('Designer'), cell('Active')] },
      { cells: [cell('Carol'), cell('Manager'), cell('On Leave')] },
    ],
  });

  builder.addBlock(emptyParagraph());

  // ---- Table 2: Column spanning ----
  builder.addBlock(paragraph([{ text: 'Table 2: Column Spanning (gridSpan)', rPr: {} }], { pStyle: 'Heading1' }));

  const col4 = Math.round(8406 / 4);
  builder.addTable({
    tblPr: {
      tblW: { w: 8406, type: 'dxa' },
      tblBorders: {
        top:     { val: 'single', sz: 4, space: 0, color: '4472C4' },
        bottom:  { val: 'single', sz: 4, space: 0, color: '4472C4' },
        start:   { val: 'single', sz: 4, space: 0, color: '4472C4' },
        end:     { val: 'single', sz: 4, space: 0, color: '4472C4' },
        insideH: { val: 'single', sz: 4, space: 0, color: '4472C4' },
        insideV: { val: 'single', sz: 4, space: 0, color: '4472C4' },
      },
    },
    colWidths: [col4, col4, col4, col4],
    rows: [
      // Merged header spanning all 4 columns
      {
        cells: [
          { tcPr: { gridSpan: 4, shd: { val: 'clear', fill: '4472C4' }, tcW: { w: 8406, type: 'dxa' as const } },
            content: [paragraph([{ text: 'Quarterly Results (Merged Header)', rPr: { b: true, color: { val: 'FFFFFF' } } }], { jc: 'center' })] },
        ],
      },
      // Sub-headers
      {
        cells: [
          cell('Q1', { bold: true, shade: 'D9E1F2' }),
          cell('Q2', { bold: true, shade: 'D9E1F2' }),
          cell('Q3', { bold: true, shade: 'D9E1F2' }),
          cell('Q4', { bold: true, shade: 'D9E1F2' }),
        ],
      },
      // Data
      { cells: [cell('100'), cell('120'), cell('115'), cell('140')] },
      // Merged last two columns
      {
        cells: [
          cell('H1 Total: 220', { gridSpan: 2, shade: 'EBF3FB' }),
          cell('H2 Total: 255', { gridSpan: 2, shade: 'EBF3FB' }),
        ],
      },
    ],
  });

  builder.addBlock(emptyParagraph());

  // ---- Table 3: Auto-width table ----
  builder.addBlock(paragraph([{ text: 'Table 3: Auto-Width Table', rPr: {} }], { pStyle: 'Heading1' }));

  builder.addTable({
    tblPr: {
      tblW: { w: 0, type: 'auto' },
      tblLayout: 'autofit',
    },
    colWidths: [2000, 4000, 2000],
    rows: [
      { cells: [cell('ID', { bold: true }), cell('Description', { bold: true }), cell('Value', { bold: true })] },
      { cells: [cell('001'), cell('First item with a longer description'), cell('$10.00')] },
      { cells: [cell('002'), cell('Second item'), cell('$25.50')] },
      { cells: [cell('003'), cell('Third item with the longest description in the table'), cell('$7.99')] },
    ],
  });

  builder.addBlock(emptyParagraph());

  // ---- Section: landscape page ----
  // Add a paragraph that ends the portrait section and starts landscape
  // Note: full header/footer support is DEFERRED_V1_1.
  // We demonstrate section type change via sectPr on the body.
  builder.addBlock(paragraph([{ text: 'Section Properties', rPr: {} }], { pStyle: 'Heading1' }));
  builder.addBlock(paragraph(['The section properties below set A4 landscape orientation for this document.']));
  builder.addBlock(paragraph([
    'Page size: w=16838 twips (A4 height), h=11906 twips (A4 width) with orient=landscape.',
  ]));

  // Override section to landscape A4
  builder.setPageSize({ w: 16838, h: 11906, orient: 'landscape' });
  builder.setPageMargins({ top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 });

  builder.addBlock(emptyParagraph());

  const buf = await builder.build();
  const out = path.join(__dirname, 'output-tables.docx');
  fs.writeFileSync(out, buf);
  console.log(`Written: ${out} (${buf.length} bytes)`);
}

main().catch(err => { console.error(err); process.exit(1); });
