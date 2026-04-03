# API Stability Classification

## Stable Public API (v1.0)

### Core Builder
- `DocxBuilder` class - All methods stable
- `DocxBuilder` constructor - Stable
- `DocxBuilder.build()` - Stable

### Factory Functions
- `paragraph(runs, pPr?)` - Stable
- `textRun(text, rPr?)` - Stable
- `emptyParagraph(pPr?)` - Stable
- `pageBreak()` - Stable

### Measurement Helpers
- `ptToTwips(pt)` - Stable
- `inToTwips(inches)` - Stable
- `cmToTwips(cm)` - Stable
- `ptToHalfPt(pt)` - Stable

### Numbering & Lists
- `bulletListLevel(options)` - Stable
- `decimalListLevel(options)` - Stable
- `arabicAbjadListLevel(options)` - Stable

### Fields & References
- `pageNumberField()` - Stable
- `totalPagesField()` - Stable
- `tocField(options)` - Stable
- `seqField(options)` - Stable
- `styleRefField(styleId, type)` - Stable
- `refField(ref)` - Stable

### Advanced Content
- `createHyperlink(url, content)` - Stable
- `createTableOfContents(options)` - Stable
- `createCaptionParagraph(label, text)` - Stable

### Type Definitions
All `Wml*` and `Docx*` type definitions are stable public API.

## Provisional API (May change)

None in v1.0 - all APIs are stable.

## Internal API (Do not use)

The following are exported for technical reasons but should not be used directly:

### Serialization Functions
- `serializeDocument(doc)` - Internal use only
- `serializeStyles(styles)` - Internal use only
- `serializeNumbering(numbering)` - Internal use only
- `serializeSettings(settings)` - Internal use only
- `serializeHeader(header)` - Internal use only
- `serializeFooter(footer)` - Internal use only
- `serializeFootnotes(footnotes)` - Internal use only
- `serializeEndnotes(endnotes)` - Internal use only

### XML Utilities
- `el(tag, attrs, ...children)` - Internal use only
- `xmlDocument(root)` - Internal use only

### Core Infrastructure
- `OpcPackage` - Internal use only
- `ContentTypeManager` - Internal use only
- `RelationshipManager` - Internal use only

## Migration Guide

### From v0.x to v1.0
- All public APIs are preserved
- Some internal method signatures may have changed but public interface is stable
- TypeScript provides compile-time safety for API usage

### Future Compatibility
- All stable APIs will maintain backward compatibility
- New features will be additive
- Breaking changes will only occur in major version bumps