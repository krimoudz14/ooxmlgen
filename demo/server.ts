/**
 * DEMO SERVER — temporary, delete demo/ to remove entirely.
 *
 * Run: npx ts-node --project tsconfig.test.json demo/server.ts
 * Open: http://localhost:3456
 */

import * as http from 'http';
import * as fs from 'fs';
import * as path from 'path';
import {
  DocxBuilder, paragraph, pageBreak, textRun,
  ptToHalfPt, ptToTwips, pageNumberField,
} from '../src/index';
import type { WmlHeaderFooterRef, WmlRunProperties } from '../src/index';

const PORT = 3456;
const HTML_FILE = path.join(__dirname, 'index.html');

function readBody(req: http.IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => { data += chunk; });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

// ---------------------------------------------------------------------------
// DOCX generation
// ---------------------------------------------------------------------------
async function generateDocx(fields: {
  title: string;
  heading: string;
  body: string;
  footnote: string;
  fontFamily: string;
  fontSize: number;       // pt
  rtl: boolean;
  pageBreak: boolean;
  headerType: string;     // 'none' | 'default' | 'firstDiff' | 'evenOdd'
  pageNumbering: boolean;
}): Promise<Buffer> {
  const builder = new DocxBuilder();

  // Core properties
  builder.setCoreProperties({
    title: fields.title || 'Untitled',
    creator: 'ooxmlgen demo',
    created: new Date(),
    modified: new Date(),
  });

  // RTL document defaults
  if (fields.rtl) {
    builder.configureDocDefaults({
      rPrDefault: { rtl: true, rFonts: { cs: fields.fontFamily, ascii: fields.fontFamily, hAnsi: fields.fontFamily } },
      pPrDefault: { bidi: true, jc: 'end' },
    });
  }

  // Styles
  const bodyRPr: WmlRunProperties = {
    sz: ptToHalfPt(fields.fontSize),
    rFonts: { ascii: fields.fontFamily, hAnsi: fields.fontFamily, cs: fields.fontFamily },
    ...(fields.rtl && { rtl: true }),
  };

  builder.addStyle({
    type: 'paragraph', styleId: 'Normal', name: 'Normal', default: true,
    pPr: {
      spacing: { after: ptToTwips(8), line: 276, lineRule: 'auto' },
      ...(fields.rtl && { bidi: true, jc: 'end' as const }),
    },
    rPr: bodyRPr,
  });

  builder.addStyle({
    type: 'paragraph', styleId: 'Heading1', name: 'heading 1',
    basedOn: 'Normal', next: 'Normal',
    pPr: {
      outlineLvl: 0, keepNext: true,
      spacing: { before: ptToTwips(12), after: ptToTwips(6) },
      ...(fields.rtl && { bidi: true, jc: 'end' as const }),
    },
    rPr: {
      b: true,
      sz: ptToHalfPt(fields.fontSize + 5),
      rFonts: { ascii: fields.fontFamily, hAnsi: fields.fontFamily, cs: fields.fontFamily },
      color: { val: '2E74B5' },
      ...(fields.rtl && { rtl: true }),
    },
  });

  // Settings
  const needsEvenOdd = fields.headerType === 'evenOdd';
  const needsFirstDiff = fields.headerType === 'firstDiff';
  if (fields.pageNumbering || needsEvenOdd) {
    builder.configureSettings({
      updateFields: fields.pageNumbering,
      evenAndOddHeaders: needsEvenOdd,
    });
  }

  // Header / Footer
  const headerRefs: WmlHeaderFooterRef[] = [];
  const footerRefs: WmlHeaderFooterRef[] = [];
  const titleText = fields.title || 'Document';

  if (fields.headerType === 'default') {
    headerRefs.push(builder.addHeader('default', [
      paragraph([titleText], { jc: 'center' }),
    ]));
  } else if (fields.headerType === 'firstDiff') {
    // First page: blank header
    headerRefs.push(builder.addHeader('first', [
      paragraph(['']),
    ]));
    // Default (all other pages): title
    headerRefs.push(builder.addHeader('default', [
      paragraph([titleText], { jc: 'center' }),
    ]));
  } else if (fields.headerType === 'evenOdd') {
    // Odd pages: title left-aligned
    headerRefs.push(builder.addHeader('default', [
      paragraph([titleText], { jc: 'start' }),
    ]));
    // Even pages: title right-aligned
    headerRefs.push(builder.addHeader('even', [
      paragraph([titleText], { jc: 'end' }),
    ]));
  }

  if (fields.pageNumbering) {
    if (fields.headerType === 'firstDiff') {
      // First page: blank footer
      footerRefs.push(builder.addFooter('first', [paragraph([''])]));
    }
    footerRefs.push(builder.addFooter('default', [
      paragraph([...pageNumberField()], { jc: 'center' }),
    ]));
    if (fields.headerType === 'evenOdd') {
      footerRefs.push(builder.addFooter('even', [
        paragraph([...pageNumberField()], { jc: 'center' }),
      ]));
    }
  }

  // Title paragraph
  if (fields.title) {
    builder.addBlock(paragraph(
      [{ text: fields.title, rPr: {} }],
      { pStyle: 'Heading1' },
    ));
  }

  // Optional heading
  if (fields.heading.trim()) {
    builder.addBlock(paragraph(
      [{ text: fields.heading, rPr: { b: true, sz: ptToHalfPt(fields.fontSize + 2) } }],
      { spacing: { before: ptToTwips(6), after: ptToTwips(4) } },
    ));
  }

  // Optional page break
  if (fields.pageBreak) {
    builder.addBlock(pageBreak());
  }

  // Body text — each line is a paragraph; footnote attached to first non-empty line
  const lines = fields.body.split('\n');
  let footnoteAttached = false;
  for (const line of lines) {
    if (fields.footnote.trim() && !footnoteAttached && line.trim()) {
      // Attach footnote to the first non-empty body paragraph
      const noteRun = builder.addFootnote([
        paragraph([fields.footnote.trim()]),
      ]);
      builder.addBlock(paragraph(
        [textRun(line, bodyRPr), noteRun],
        { ...(fields.rtl && { bidi: true, jc: 'end' as const }) },
      ));
      footnoteAttached = true;
    } else {
      builder.addBlock(paragraph(
        [textRun(line, bodyRPr)],
        { ...(fields.rtl && { bidi: true, jc: 'end' as const }) },
      ));
    }
  }

  // Section properties — wire everything together
  builder.setSectionProperties({
    pgSz:  { w: 11906, h: 16838 },
    pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
    ...(fields.rtl && { bidi: true, rtlGutter: true }),
    ...(needsFirstDiff && { titlePg: true }),
    ...(headerRefs.length > 0 && { headerReference: headerRefs }),
    ...(footerRefs.length > 0 && { footerReference: footerRefs }),
  });

  return builder.build();
}

// ---------------------------------------------------------------------------
// HTTP server
// ---------------------------------------------------------------------------
const server = http.createServer(async (req, res) => {
  if (req.method === 'GET' && req.url === '/') {
    const html = fs.readFileSync(HTML_FILE, 'utf8');
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(html);
    return;
  }

  if (req.method === 'POST' && req.url === '/generate') {
    try {
      const raw = await readBody(req);
      const params = new URLSearchParams(raw);

      const fields = {
        title:         params.get('title')         ?? '',
        heading:       params.get('heading')       ?? '',
        body:          params.get('body')          ?? '',
        footnote:      params.get('footnote')      ?? '',
        fontFamily:    params.get('fontFamily')    || 'Calibri',
        fontSize:      parseInt(params.get('fontSize') || '11', 10),
        rtl:           params.get('rtl')           === 'on',
        pageBreak:     params.get('pageBreak')     === 'on',
        headerType:    params.get('headerType')    || 'none',
        pageNumbering: params.get('pageNumbering') === 'on',
      };

      const buf = await generateDocx(fields);
      const filename = (fields.title.replace(/[^a-z0-9]/gi, '_') || 'document') + '.docx';

      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Content-Length': buf.length,
      });
      res.end(buf);
    } catch (err) {
      console.error(err);
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.end('Error: ' + String(err));
    }
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`Demo running at http://localhost:${PORT}`);
  console.log('Press Ctrl+C to stop.');
});
