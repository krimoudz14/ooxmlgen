# ooxmlgen

TypeScript library for generating DOCX / WordprocessingML documents in Node.js.

Built directly from the ECMA-376 5th edition schema. Produces valid `.docx` files compatible with Microsoft Word.

---

## Installation

```bash
npm install ooxmlgen
```

---

## Quick Start

```typescript
import { DocxBuilder, paragraph } from 'ooxmlgen';
import * as fs from 'fs';

const builder = new DocxBuilder();

builder.setCoreProperties({ title: 'My Document', creator: 'My App' });
builder.setPageSize({ w: 11906, h: 16838 }); // A4
builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800,
                         header: 720, footer: 720, gutter: 0 });

builder.addBlock(paragraph(['Hello, World!']));

const buffer = await builder.build();
fs.writeFileSync('output.docx', buffer);
```

---

## Features

- **Paragraphs & runs** — text, bold, italic, underline, color, font, size, RTL
- **Tables** — borders, cell merging, column widths, cell margins
- **Styles** — paragraph and character style definitions with inheritance
- **Numbering** — bulleted, decimal, arabicAbjad lists
- **Headers / footers** — default, first-page, even/odd variants
- **Page number fields** — PAGE, NUMPAGES
- **Footnotes / endnotes** — with required separator entries
- **Hyperlinks** — external URLs and internal anchors
- **TOC / SEQ fields** — table of contents, caption sequences
- **Document defaults** — docDefaults for consistent base formatting
- **RTL / Arabic** — bidi, textDirection, complex script fonts, rtlGutter
- **Document settings** — updateFields, evenAndOddHeaders, themeFontLang
- **Core / extended properties** — title, creator, application, etc.

---

## API

### DocxBuilder

```typescript
const builder = new DocxBuilder();

// Page setup
builder.setPageSize({ w: 11906, h: 16838 });
builder.setPageMargins({ top: 1440, right: 1800, bottom: 1440, left: 1800,
                         header: 720, footer: 720, gutter: 0 });
builder.setSectionProperties({ pgSz: ..., pgMar: ..., headerReference: [...] });

// Properties
builder.setCoreProperties({ title, creator, created, modified });
builder.setExtendedProperties({ application, appVersion });

// Styles & defaults
builder.addStyle({ type: 'paragraph', styleId: 'Normal', name: 'Normal', ... });
builder.configureDocDefaults({ rPrDefault: { rtl: true, rFonts: { cs: 'Arial' } } });

// Numbering
const numId = builder.defineNumbering({ levels: [...] });

// Settings
builder.configureSettings({ updateFields: true, evenAndOddHeaders: true });

// Headers / footers
const hdrRef = builder.addHeader('default', [paragraph(['My Header'])]);
const ftrRef = builder.addFooter('default', [paragraph([...pageNumberField()])]);

// Footnotes
const noteRun = builder.addFootnote([paragraph(['Footnote text.'])]);

// Hyperlinks
const rId = builder.addHyperlinkRelationship('https://example.com');

// Content
builder.addBlock(paragraph(['text']));
builder.addBlock(paragraph([textRun('bold', { b: true })]));
builder.addTable({ colWidths: [3000, 3000], rows: [...] });

// Build
const buffer = await builder.build(); // → Buffer
```

### Factory helpers

```typescript
import { paragraph, textRun, emptyParagraph, pageBreak,
         pageNumberField, totalPagesField, pageXofYRuns,
         tocField, seqField, fieldRuns,
         createHyperlink, createCaptionParagraph, createTableOfContents,
         bulletListLevel, decimalListLevel, arabicAbjadListLevel } from 'ooxmlgen';
```

### Measurement helpers

```typescript
import { ptToTwips, inToTwips, cmToTwips, ptToHalfPt } from 'ooxmlgen';

ptToTwips(12)    // 12pt → 240 twips
inToTwips(1)     // 1 inch → 1440 twips
cmToTwips(2.54)  // 2.54cm → 1440 twips
ptToHalfPt(11)   // 11pt → 22 half-points (for w:sz)
```

---

## Examples

See the [`examples/`](examples/) directory for working scripts:

```bash
npx ts-node --project tsconfig.test.json examples/generate.ts
npx ts-node --project tsconfig.test.json examples/example-batch1.ts
npx ts-node --project tsconfig.test.json examples/example-batch5.ts
```

Or use the `generate` script:

```bash
npm run generate
```

---

## Development

```bash
npm install
npm run build    # compile TypeScript → dist/
npm test         # run 246 tests
```

---

## Generated Package Structure

Parts emitted by ooxmlgen (only when the feature is used):

```
[Content_Types].xml
_rels/.rels
word/document.xml
word/styles.xml
word/_rels/document.xml.rels
docProps/core.xml
docProps/app.xml
word/numbering.xml       # when defineNumbering() is called
word/settings.xml        # when configureSettings() is called
word/footnotes.xml       # when addFootnote() is called
word/endnotes.xml        # when addEndnote() is called
word/header*.xml         # when addHeader() is called
word/footer*.xml         # when addFooter() is called
```

---

## Deferred

- Images / drawings → v1.1
- Font table → v1.1
- Comments, track changes → v2
- XLSX / SpreadsheetML → v2
- PPTX / PresentationML → v3

---

## Source Traceability

Every TypeScript construct is traced to its ECMA-376 5th edition source in [`SPEC_TRACEABILITY.md`](SPEC_TRACEABILITY.md).

**246 automated tests · ECMA-376 5th edition · Node.js · TypeScript**
