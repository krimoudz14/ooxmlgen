# Advanced Word Mechanisms — Package Design

**Date:** 2026-04-01  
**Purpose:** Define the reusable mechanism API surface for the next package version.  
**Principle:** The package exposes mechanisms. Applications compose them into documents.

---

## Design Principle

```
Package layer (this file):
  - numbering definitions
  - header/footer parts
  - page number fields
  - footnotes/endnotes
  - field infrastructure
  - TOC/caption/generated-list scaffolding
  - RTL/Arabic formatting primitives
  - document settings

Application layer (NOT in this package):
  - thesis cover page
  - chapter structure
  - university branding
  - fixed section order
  - bibliography page layout
```

---

## Reusable Feature Map

### A. Section and Page Mechanics

**Already implemented (v1):**
- `WmlSectionProperties` — type, pgSz, pgMar, orient, bidi, titlePg, vAlign
- Section break types: nextPage, continuous, evenPage, oddPage, nextColumn

**Missing — to add:**
- `pgNumType` on sectPr — page number format (decimal, upperRoman, arabicAbjad, etc.) and start value
- `headerReference` / `footerReference` on sectPr — link sections to header/footer parts
- `textDirection` on sectPr — explicit section text direction
- Mid-document section breaks (sectPr inside paragraph pPr, not just final body sectPr)
- `rtlGutter` on sectPr

**Proposed API additions:**
```typescript
// On WmlSectionProperties:
pgNumType?: { fmt?: WmlNumberFormat; start?: number }
headerReference?: WmlHeaderFooterRef[]   // up to 3: default, first, even
footerReference?: WmlFooterRef[]
textDirection?: WmlTextDirectionValue
rtlGutter?: boolean

// Mid-document section break:
// A paragraph's pPr.sectPr triggers a section break before the next paragraph
// Already in CT_PPr — needs to be wired in WmlParagraphProperties
sectPr?: WmlSectionProperties  // already in CT_PPr, not yet in our type
```

---

### B. Header / Footer Mechanics

**Missing — to add:**
- `word/header1.xml` (default), `word/header2.xml` (even), `word/header3.xml` (first page)
- `word/footer1.xml`, `word/footer2.xml`, `word/footer3.xml`
- Relationship entries in `word/_rels/document.xml.rels`
- Content type entries in `[Content_Types].xml`
- Header/footer content model: same as body (EG_BlockLevelElts)

**Proposed API:**
```typescript
// Builder methods:
builder.addHeader(type: 'default' | 'first' | 'even', content: WmlBlockElement[]): string  // returns rId
builder.addFooter(type: 'default' | 'first' | 'even', content: WmlBlockElement[]): string  // returns rId

// Section references these by rId:
sectPr.headerReference = [{ type: 'default', id: rId }]
sectPr.footerReference = [{ type: 'default', id: rId }]
```

---

### C. Page Number Fields

**Missing — to add:**
- `w:fldChar` with `fldCharType="begin"/"separate"/"end"` in run content
- `w:instrText` with field instruction string (e.g. `PAGE`, `NUMPAGES`, `TOC`, `SEQ`)
- `w:t` as display text between separate and end fldChar

**Proposed API:**
```typescript
// Factory helpers:
pageNumberField(): WmlRun[]          // emits PAGE field sequence
totalPagesField(): WmlRun[]          // emits NUMPAGES field sequence
pageXofYRuns(): WmlRun[]             // emits "PAGE / NUMPAGES" sequence

// Low-level:
fieldRuns(instruction: string, displayText?: string): WmlRun[]
```

**Field character sequence (from CT_FldChar spec):**
```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>   <!-- cached display value -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

---

### D. Footnote / Endnote Mechanics

**Missing — to add:**
- `word/footnotes.xml` part with CT_Footnotes
- `word/endnotes.xml` part with CT_Endnotes
- `w:footnoteReference` in run content (CT_FtnEdnRef)
- `w:endnoteReference` in run content
- Footnote ID management (auto-increment)
- Footnote content: same block-level model as body
- Separator and continuation separator footnotes (required by spec)
- `footnotePr`/`endnotePr` in sectPr and settings

**Proposed API:**
```typescript
// Builder method:
builder.addFootnote(content: WmlBlockElement[]): WmlRun  // returns footnoteReference run
builder.addEndnote(content: WmlBlockElement[]): WmlRun   // returns endnoteReference run

// Configuration:
builder.configureFootnotes(props: FootnoteConfig): this
// FootnoteConfig: { numFmt?, startNum?, restartEach?: 'page' | 'section' | 'continuous' }
```

---

### E. Field Infrastructure

**Missing — to add:**
- `WmlFldChar` run content type (fldCharType: begin/separate/end)
- `WmlInstrText` run content type
- These extend `WmlRunContent` union
- Factory helpers for common fields

**Proposed run content additions:**
```typescript
interface WmlFldChar {
  _type: 'fldChar';
  fldCharType: 'begin' | 'separate' | 'end';
  fldLock?: boolean;
  dirty?: boolean;
}

interface WmlInstrText {
  _type: 'instrText';
  text: string;
  space?: 'preserve' | 'default';
}
```

---

### F. TOC and Generated Lists

**Missing — to add:**
- TOC field: `TOC \o "1-3" \h \z \u` instruction
- SEQ field: `SEQ Figure \* ARABIC` for figure/table numbering
- TC field: `TC "text" \f F` for table of contents entries
- `updateFields` in settings (required for Word to build TOC on open)
- Caption paragraph style + SEQ field combination

**Proposed API:**
```typescript
// Factory helpers:
tableOfContents(options?: { levels?: string; hyperlinks?: boolean }): WmlParagraph[]
captionParagraph(label: string, number: string | WmlRun[], text: string): WmlParagraph
listOfFigures(): WmlParagraph[]
listOfTables(): WmlParagraph[]

// These produce the correct field markup — applications compose them into pages
```

---

### G. Numbering Definitions

**Missing — to add:**
- `word/numbering.xml` part with CT_Numbering
- CT_AbstractNum with CT_Lvl definitions
- CT_Num referencing abstractNum
- `w:numPr` serialization in paragraph pPr
- Relationship entry in `word/_rels/document.xml.rels`

**Proposed API:**
```typescript
// Builder method:
builder.defineNumbering(def: NumberingDefinition): number  // returns numId

interface NumberingDefinition {
  abstractNumId?: number;  // auto-assigned if omitted
  multiLevelType?: 'singleLevel' | 'multilevel' | 'hybridMultilevel';
  levels: NumberingLevel[];
}

interface NumberingLevel {
  ilvl: number;           // 0-8
  start?: number;         // default 1
  numFmt?: WmlNumberFormat;  // decimal, upperRoman, bullet, arabicAbjad, etc.
  lvlText?: string;       // e.g. "%1." or "•"
  jc?: WmlJustificationValue;
  pPr?: WmlParagraphProperties;
  rPr?: WmlRunProperties;
  pStyle?: string;        // link to paragraph style
}

// Paragraph usage:
pPr.numPr = { numId: 1, ilvl: 0 }  // already in type, now serialized
```

---

### H. RTL / Arabic Formatting Primitives

**Already working:**
- `w:bidi` on paragraph (RTL paragraph direction)
- `w:rtl` on run (RTL run direction)
- `rFonts.cs` / `rFonts.eastAsia` (Arabic/complex script fonts)
- `szCs` (complex script font size)
- `bCs`, `iCs` (bold/italic for complex script)
- `lang.bidi` (Arabic language tag on run)

**Missing — to add:**
- `w:textDirection` on paragraph (explicit lr/rl/tb)
- `w:textDirection` on section
- `w:bidi` on section (already in WmlSectionProperties but not wired to settings)
- `w:docDefaults` in styles.xml — set Arabic font as document default
- `word/settings.xml` with `themeFontLang` for Arabic
- `w:bidi` in CT_Settings (document-level RTL)

**Proposed API additions:**
```typescript
// On WmlParagraphProperties (already has bidi):
textDirection?: WmlTextDirectionValue  // 'tb' | 'rl' | 'lr' | 'tbV' | 'rlV' | 'lrV'

// On WmlSectionProperties (already has bidi):
textDirection?: WmlTextDirectionValue

// On DocxBuilder:
builder.setDocumentDefaults(defaults: DocumentDefaults): this
// DocumentDefaults: { rPr?: WmlRunProperties; pPr?: WmlParagraphProperties }

builder.configureSettings(settings: DocxSettings): this
// DocxSettings: { evenAndOddHeaders?, updateFields?, bidi?, themeFontLang?, ... }
```

**Arabic document pattern (application-level, not package-level):**
```typescript
// An Arabic thesis app would do:
builder.configureSettings({ bidi: true, updateFields: true })
builder.setDocumentDefaults({
  rPr: { rFonts: { cs: 'Times New Roman', ascii: 'Times New Roman' }, szCs: 24, lang: { bidi: 'ar-SA' } },
  pPr: { bidi: true, jc: 'both' }
})
// Then add content with RTL paragraphs
```

---

### I. Reference / Citation Helpers

**Scope:** Reusable helpers for note insertion and reference formatting.  
**Not in scope:** Hardcoded bibliography sections or citation styles.

**Proposed API:**
```typescript
// Footnote-based citation (uses footnote mechanism from D):
builder.addFootnote([paragraph(['Author, Title, Year.'])])  // returns footnoteReference run

// Endnote-based citation:
builder.addEndnote([paragraph(['Author, Title, Year.'])])   // returns endnoteReference run

// Reference paragraph style (application defines, package provides mechanism):
// Application adds a 'References' style and uses it for reference paragraphs
// Package provides no hardcoded reference section
```

---

## ECMA Source Files to Reopen

| File | Size | Current status | Needed for |
|---|---|---|---|
| `5TH EDITION PART1/wml.xsd` | 163KB | Partially read | CT_TcPr full, CT_Numbering, CT_FtnEdn, CT_Settings, CT_FldChar, CT_Hyperlink, CT_PageNumber, CT_TextDirection — **all now read** |
| `5TH EDITION PART2/opc-contentTypes.xsd` | 2KB | Fully read | Header/footer/footnote/endnote/numbering/settings content types |
| `5TH EDITION PART2/opc-relationships.xsd` | 1.4KB | Fully read | New relationship types for new parts |

No additional schema files need to be opened. All required types are in `wml.xsd` which has now been read for all needed sections.

---

## Implementation Batch Plan

### Batch 1 — Foundation (section mechanics, numbering, settings, table cell completion)
1. `pgNumType` on sectPr (CT_PageNumber — read)
2. `textDirection` on sectPr (CT_TextDirection — read)
3. Mid-document section breaks via `pPr.sectPr`
4. Document settings part (`word/settings.xml`) — minimal: `evenAndOddHeaders`, `updateFields`, `bidi`, `themeFontLang`
5. Numbering definitions (`word/numbering.xml`) — CT_AbstractNum, CT_Lvl, CT_Num
6. `w:numPr` serialization in paragraph pPr
7. Table cell borders (`w:tcBorders`) and margins (`w:tcMar`) — complete CT_TcPr

### Batch 2 — Headers, Footers, Page Numbers
1. Header parts (`word/header*.xml`) — default, first, even
2. Footer parts (`word/footer*.xml`) — default, first, even
3. `headerReference`/`footerReference` in sectPr
4. Field infrastructure: `WmlFldChar`, `WmlInstrText` run content types
5. Page number field helpers: `pageNumberField()`, `totalPagesField()`

### Batch 3 — Footnotes and Endnotes
1. `word/footnotes.xml` with separator/continuation separator
2. `word/endnotes.xml`
3. `w:footnoteReference` / `w:endnoteReference` in run content
4. Footnote/endnote ID management
5. `footnotePr`/`endnotePr` in sectPr and settings

### Batch 4 — Field Automation (TOC, Captions, Generated Lists, Hyperlinks)
1. Hyperlinks (`w:hyperlink` in paragraph content)
2. TOC field helper
3. Caption paragraph helper (SEQ field)
4. List of Figures / Tables helpers
5. `updateFields` in settings (required for TOC)

### Batch 5 — RTL/Arabic Refinements and Document Defaults
1. `w:textDirection` on paragraph and section
2. `w:docDefaults` in styles.xml
3. `configureSettings()` builder method
4. `setDocumentDefaults()` builder method
5. Arabic document pattern validation
6. Font table (`word/fontTable.xml`) — optional but recommended for Arabic

---

## Public API Naming Conventions

All new APIs must be reusable and generic:

| ✓ Correct (mechanism) | ✗ Wrong (template-specific) |
|---|---|
| `addHeader(type, content)` | `addThesisHeader()` |
| `addFootnote(content)` | `addCitationFootnote()` |
| `defineNumbering(def)` | `addChapterNumbering()` |
| `configureSettings(opts)` | `setArabicMode()` |
| `tableOfContents(opts)` | `addThesisTOC()` |
| `captionParagraph(label, num, text)` | `addFigureCaption()` |
| `pageNumberField()` | `addPageNumber()` |
| `setDocumentDefaults(defaults)` | `setArabicDefaults()` |

---

## Artifact Update Plan

After each batch:
1. Update `repo-manifest.json` — mark newly read wml.xsd sections
2. Update `spec-extraction.json` — add extraction records for new CT_* types
3. Update `SPEC_TRACEABILITY.md` — add traceability entries for new constructs
4. Add automated tests
5. Generate example DOCX and inspect with unzip
