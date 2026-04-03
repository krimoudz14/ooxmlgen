# Migration Guide & Examples

This guide helps you migrate to ooxmlgen v1.0 and provides common usage patterns.

## Migration from Other Libraries

### From docx (npm package)

```typescript
// Before (docx library)
import { Document, Paragraph, TextRun } from 'docx';

const doc = new Document({
  sections: [{
    properties: {},
    children: [
      new Paragraph({
        children: [new TextRun("Hello World")]
      })
    ]
  }]
});

// After (ooxmlgen)
import { DocxBuilder, paragraph, textRun } from 'ooxmlgen';

const builder = new DocxBuilder();
builder.addBlock(paragraph([textRun("Hello World")]));
const buffer = await builder.build();
```

### From officegen

```typescript
// Before (officegen)
const officegen = require('officegen');
const docx = officegen('docx');
docx.on('finalize', (written) => console.log('Done'));
docx.on('error', (err) => console.error(err));

// After (ooxmlgen)
import { DocxBuilder, paragraph } from 'ooxmlgen';

const builder = new DocxBuilder();
builder.addBlock(paragraph(['Hello World']));
const buffer = await builder.build();
// Write buffer to file
```

## Common Patterns

### Business Document Template

```typescript
const builder = new DocxBuilder();

// Document setup
builder
  .setCoreProperties({
    title: 'Business Report',
    creator: 'Your App',
    created: new Date()
  })
  .setPageSize({ w: ptToTwips(8.5), h: ptToTwips(11) })
  .setPageMargins({ top: 1440, right: 1440, bottom: 1440, left: 1440 });

// Styles
builder.addStyle({
  type: 'paragraph',
  styleId: 'Heading1',
  name: 'heading 1',
  pPr: { outlineLvl: 0, keepNext: true },
  rPr: { b: true, sz: ptToHalfPt(16) }
});

// Header
builder.addHeader([
  paragraph(['Company Name'], { jc: 'center' })
], 'default');

// Content
builder.addBlock(paragraph(['Executive Summary'], { pStyle: 'Heading1' }));
builder.addBlock(paragraph(['This is the executive summary content.']));

// Footer with page numbers
builder.addFooter([
  paragraph([pageNumberField()], { jc: 'center' })
], 'default');

const docx = await builder.build();
```

### Academic Document with Footnotes

```typescript
const builder = new DocxBuilder();

// Define a footnote
const footnoteRef = builder.defineFootnote('fn1', [
  paragraph(['This is a footnote explaining the source.'])
]);

// Use in document
builder.addBlock(paragraph([
  'This statement requires a citation',
  footnoteRef
]));
```

### Arabic/Russian Document

```typescript
const builder = new DocxBuilder();

// Arabic document defaults
builder.configureDocDefaults({
  rPrDefault: {
    rtl: true,
    rFonts: { cs: 'Arial Unicode MS', eastAsia: 'MS Gothic' }
  },
  pPrDefault: {
    bidi: true,
    jc: 'right',
    textDirection: 'rl'
  }
});

// Arabic content
builder.addBlock(paragraph(['العنوان الرئيسي'], {
  bidi: true,
  jc: 'right'
}));
```

## Integration Examples

### Express.js API Endpoint

```typescript
import express from 'express';
import { DocxBuilder, paragraph } from 'ooxmlgen';

const app = express();

app.get('/api/document', async (req, res) => {
  const builder = new DocxBuilder();
  builder.addBlock(paragraph([`Generated at ${new Date().toISOString()}`]));

  const buffer = await builder.build();

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="document.docx"');
  res.send(buffer);
});
```

### React Component (Client-side)

```typescript
import React, { useState } from 'react';
import { DocxBuilder, paragraph } from 'ooxmlgen';

function DocumentGenerator() {
  const [isGenerating, setIsGenerating] = useState(false);

  const generateDocument = async () => {
    setIsGenerating(true);
    try {
      const builder = new DocxBuilder();
      builder.addBlock(paragraph(['Hello from React!']));

      const buffer = await builder.build();

      // Download the file
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.docx';
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <button onClick={generateDocument} disabled={isGenerating}>
      {isGenerating ? 'Generating...' : 'Generate DOCX'}
    </button>
  );
}
```

## Best Practices

1. **Use TypeScript** - Full type safety prevents runtime errors
2. **Set document properties** - Always call `setCoreProperties()` for metadata
3. **Configure settings** - Use `configureSettings({ updateFields: true })` for dynamic content
4. **Define styles first** - Add styles before using them in content
5. **Test with Word** - Always verify generated documents in Microsoft Word
6. **Use measurement helpers** - Convert between units properly with `ptToTwips()`, etc.

## Troubleshooting

### Document won't open
- Ensure all required parts are generated (check with unzip)
- Verify XML is well-formed
- Test with minimal document first

### Formatting not applied
- Check style IDs match exactly (case-sensitive)
- Verify properties are set on correct elements
- Use `docDefaults` for consistent formatting

### Arabic text not displaying
- Use `rtl: true` on runs with Arabic text
- Set `rFonts: { cs: 'Arial Unicode MS' }` for complex scripts
- Configure `themeFontLang` in settings