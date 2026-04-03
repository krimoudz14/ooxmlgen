/**
 * Batch 5 example — Document Defaults and RTL/Arabic Support
 *
 * Demonstrates:
 * - configureDocDefaults(): Arabic-friendly run and paragraph defaults
 * - RTL section properties (bidi, rtlGutter, textDirection)
 * - Arabic-friendly settings (bidi, themeFontLang)
 * - Mixed Arabic/English content in the same document
 * - arabicAbjad numbering for Arabic lists
 * - RTL paragraph and run properties
 *
 * Output: output/batch5-rtl-arabic.docx
 *
 * Inspect with:
 *   unzip -l output/batch5-rtl-arabic.docx
 *   unzip -p output/batch5-rtl-arabic.docx word/styles.xml
 *   unzip -p output/batch5-rtl-arabic.docx word/settings.xml
 *   unzip -p output/batch5-rtl-arabic.docx word/document.xml
 */

import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder,
  paragraph,
  textRun,
  pageBreak,
  ptToTwips,
  ptToHalfPt,
  inToTwips,
} from '../src/index';

async function main() {
  const b = new DocxBuilder();

  // ---- Document properties ----
  b.setCoreProperties({
    title: 'Batch 5 — RTL/Arabic Document Demo',
    creator: 'ooxmlgen',
  });

  // ---- Arabic document defaults ----
  // These apply to all content that doesn't have explicit formatting.
  // Source: CT_DocDefaults (wml.xsd line 2688)
  b.configureDocDefaults({
    rPrDefault: {
      rFonts: {
        ascii:    'Times New Roman',
        hAnsi:    'Times New Roman',
        cs:       'Traditional Arabic',
        eastAsia: 'Times New Roman',
      },
      lang: { bidi: 'ar-SA', val: 'en-US' },
      szCs: ptToHalfPt(14),
    },
    pPrDefault: {
      bidi: true,
      jc: 'end',
    },
  });

  // ---- Arabic-friendly settings ----
  b.configureSettings({
    bidi: true,
    themeFontLang: { bidi: 'ar-SA', val: 'en-US' },
    updateFields: true,
    defaultTabStop: 720,
  });

  // ---- Arabic numbering ----
  const arabicNumId = b.defineNumbering({
    levels: [{
      ilvl: 0,
      numFmt: 'arabicAbjad',
      lvlText: '%1.',
      lvlJc: 'start',
      suff: 'tab',
      pPr: { bidi: true, ind: { start: 720, hanging: 360 } },
    }],
  });

  // ---- RTL section ----
  b.setSectionProperties({
    pgSz: { w: 11906, h: 16838 },
    pgMar: {
      top: inToTwips(1), right: inToTwips(1.25),
      bottom: inToTwips(1), left: inToTwips(1.25),
      header: 720, footer: 720, gutter: 0,
    },
    bidi: true,
    rtlGutter: true,
  });

  // ---- Arabic content ----

  // Arabic title
  b.addBlock(paragraph(
    [{ text: 'مرحبا بالعالم', rPr: { rtl: true, cs: true, b: true, bCs: true, szCs: ptToHalfPt(18) } }],
    { bidi: true, jc: 'center', spacing: { after: 400 } },
  ));

  // Arabic paragraph
  b.addBlock(paragraph(
    [{ text: 'هذا مستند تجريبي يوضح دعم اللغة العربية في مكتبة ooxmlgen.', rPr: { rtl: true, cs: true } }],
    { bidi: true, jc: 'end', spacing: { after: 200 } },
  ));

  // Mixed Arabic/English paragraph
  b.addBlock(paragraph([
    { text: 'النص العربي: ', rPr: { rtl: true, cs: true } },
    { text: 'English text mixed in', rPr: { rtl: false } },
    { text: ' — نهاية الجملة.', rPr: { rtl: true, cs: true } },
  ], { bidi: true, jc: 'end', spacing: { after: 200 } }));

  // Arabic numbered list
  b.addBlock(paragraph(
    [{ text: 'البند الأول', rPr: { rtl: true, cs: true } }],
    { bidi: true, numPr: { numId: arabicNumId, ilvl: 0 }, spacing: { after: 100 } },
  ));
  b.addBlock(paragraph(
    [{ text: 'البند الثاني', rPr: { rtl: true, cs: true } }],
    { bidi: true, numPr: { numId: arabicNumId, ilvl: 0 }, spacing: { after: 100 } },
  ));
  b.addBlock(paragraph(
    [{ text: 'البند الثالث', rPr: { rtl: true, cs: true } }],
    { bidi: true, numPr: { numId: arabicNumId, ilvl: 0 }, spacing: { after: 400 } },
  ));

  // Page break
  b.addBlock(pageBreak());

  // English section (LTR override)
  b.addBlock(paragraph(
    [{ text: 'English Section', rPr: { b: true, sz: ptToHalfPt(16) } }],
    { bidi: false, jc: 'start', spacing: { after: 200 } },
  ));

  b.addBlock(paragraph(
    ['This paragraph demonstrates left-to-right English content in the same document.'],
    { bidi: false, jc: 'start', spacing: { after: 200 } },
  ));

  b.addBlock(paragraph([
    textRun('Bold', { b: true }),
    textRun(', '),
    textRun('italic', { i: true }),
    textRun(', and '),
    textRun('underlined', { u: 'single' }),
    textRun(' text in English.'),
  ], { bidi: false, jc: 'start' }));

  // ---- Build and write ----
  const outDir = path.join(__dirname, '..', 'output');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outPath = path.join(outDir, 'batch5-rtl-arabic.docx');
  const buf = await b.build();
  fs.writeFileSync(outPath, buf);

  console.log(`Written: ${outPath}`);
  console.log(`Size: ${buf.length} bytes`);
  console.log('');
  console.log('Inspect with:');
  console.log(`  unzip -p "${outPath}" word/styles.xml`);
  console.log(`  unzip -p "${outPath}" word/settings.xml`);
  console.log(`  unzip -p "${outPath}" word/document.xml`);
}

main().catch(err => { console.error(err); process.exit(1); });
