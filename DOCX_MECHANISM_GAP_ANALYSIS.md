# DOCX Mechanism Gap Analysis

**Date:** 2026-04-01  
**Package version:** v1 (validated)  
**Purpose:** Map missing reusable Word-generation mechanisms to ECMA source files, read status, and implementation batches.  
**Scope:** General-purpose Word engine. No thesis-specific content hardcoded here.

---

## How to read this table

- **Existing support** = what v1 already emits correctly  
- **Missing support** = what the package cannot yet produce  
- **Source read status** = whether the relevant wml.xsd / opc-*.xsd sections were actually read  
- **Batch** = recommended implementation order (see ADVANCED_WORD_MECHANISMS.md)  
- **Risk** = implementation complexity or ambiguity

---

## Gap Analysis Table

| # | Reusable mechanism | Why needed by downstream products | Evidence from use cases | Existing support | Missing support | Needed OOXML parts/files | Source read status | Batch | Risk |
|---|---|---|---|---|---|---|---|---|---|
| 1 | **Numbering definitions** | Lists (ordered, unordered, outline), heading numbering, caption numbering | Thesis: chapter numbers, section numbers, figure/table numbering; Reports: numbered sections; Any structured doc | `WmlNumberingProperties` type stub exists; serializer silently drops it | `word/numbering.xml` not emitted; CT_AbstractNum, CT_Lvl, CT_Num not serialized; `w:numPr` not emitted | `word/numbering.xml`, `word/_rels/document.xml.rels` entry | **Fully read** — CT_Lvl (line 2961), CT_AbstractNum (line 2989), CT_Num (line 3004), CT_Numbering (line 3015) | **Batch 1** | Medium — abstractNum/num ID management; level text patterns |
| 2 | **Header parts** | Page headers with document title, chapter name, page numbers | Thesis: running header with chapter title; Reports: company name in header; Any multi-page doc | `titlePg` flag in WmlSectionProperties; `headerReference` field exists in CT_SectPr but not serialized | `word/header1.xml` (and header2, header3) not emitted; `headerReference` not serialized in `w:sectPr`; no header content API | `word/header*.xml`, `word/_rels/document.xml.rels` entries | **Fully read** — CT_HdrFtr (line 1528), EG_HdrFtrReferences (line 1522), CT_HdrFtrRef (line 1514) | **Batch 2** | Medium — part naming, relationship wiring, first/even/default variants |
| 3 | **Footer parts** | Page footers with page numbers, document info | Thesis: page numbers in footer; Reports: confidentiality notice; Any paginated doc | Same as headers — `footerReference` not serialized | `word/footer*.xml` not emitted; no footer content API | `word/footer*.xml`, `word/_rels/document.xml.rels` entries | **Fully read** — same CT_HdrFtr applies to both hdr and ftr | **Batch 2** | Medium — same as headers |
| 4 | **Page number fields** | Automatic page numbering in headers/footers | Thesis: page numbers; Reports: page X of Y; Any paginated doc | None — `w:fldChar` and `w:instrText` not in run content union | Field character sequence (`w:fldChar begin`, `w:instrText PAGE`, `w:fldChar end`) not serializable; no field abstraction | Inline in `word/document.xml` or header/footer XML | **Fully read** — CT_FldChar (line 1210), ST_FldCharType (begin/separate/end), CT_SimpleField (line 1158) | **Batch 2** | Low-Medium — field char sequence is mechanical; PAGE/NUMPAGES are simple |
| 5 | **Section breaks with header/footer linkage** | Multi-section documents with different headers/footers per section | Thesis: front matter (Roman numerals) vs body (Arabic numerals); Reports: cover page vs content | `w:sectPr` type emitted; `pgSz`/`pgMar` work | `headerReference`/`footerReference` not serialized; `pgNumType` not serialized; `titlePg` not wired to actual header parts; multi-section body not supported | Inline in `word/document.xml` sectPr elements | **Fully read** — CT_SectPr (line 1566), CT_PageNumber (line 1457), EG_SectPrContents (line 1532) | **Batch 1** | Medium — sectPr in body paragraphs (not just final sectPr) for mid-document section breaks |
| 6 | **Page numbering configuration** | Control page number format and start value per section | Thesis: front matter i,ii,iii then body 1,2,3; Reports: restart at 1 after cover | None — `pgNumType` not serialized | `w:pgNumType` with `fmt` and `start` not emitted in sectPr | Inline in `w:sectPr` | **Fully read** — CT_PageNumber (line 1457), ST_NumberFormat (line 1294) including `arabicAbjad`, `decimal`, `upperRoman`, `lowerRoman` | **Batch 1** | Low — straightforward attribute serialization |
| 7 | **Footnotes** | Academic citations, legal references, explanatory notes | Thesis: citation footnotes; Academic articles: source notes; Legal docs: statutory references | `w:footnoteReference` in EG_RunInnerContent but not in WmlRunContent union | `word/footnotes.xml` not emitted; footnote reference not serializable; footnote content not serializable; footnote ID management missing | `word/footnotes.xml`, `word/_rels/document.xml.rels` entry | **Fully read** — CT_FtnEdn (line 2491), CT_FtnEdnRef (line 2484), CT_Footnotes (line 3245), CT_FtnProps (line 2506) | **Batch 3** | Medium — ID management; separator/continuation footnotes; footnote properties |
| 8 | **Endnotes** | End-of-document references | Academic articles: bibliography-style endnotes; Legal docs | Same as footnotes | `word/endnotes.xml` not emitted; endnote reference not serializable | `word/endnotes.xml`, `word/_rels/document.xml.rels` entry | **Fully read** — CT_Endnotes (line 3253), CT_EdnProps (line 2513), same CT_FtnEdn type | **Batch 3** | Low — same structure as footnotes |
| 9 | **Document settings part** | RTL document direction, even/odd headers, field update, footnote/endnote global config | Thesis: Arabic RTL document; Any doc needing `evenAndOddHeaders`, `updateFields` | None — `word/settings.xml` not emitted | `word/settings.xml` not emitted; `evenAndOddHeaders`, `updateFields`, `footnotePr`, `endnotePr`, `themeFontLang` not configurable | `word/settings.xml`, `word/_rels/document.xml.rels` entry | **Fully read** — CT_Settings (line 2733), key fields: `evenAndOddHeaders`, `updateFields`, `footnotePr`, `endnotePr`, `themeFontLang`, `bidi` | **Batch 1** | Low-Medium — large CT_Settings but only a subset needed for v2 |
| 10 | **RTL paragraph direction** | Arabic/Hebrew/Persian text paragraphs | Arabic thesis: all body paragraphs RTL; Arabic reports; Mixed Arabic/English docs | `bidi` flag exists on `WmlParagraphProperties` and `WmlSectionProperties` — **already serialized** | `w:bidi` on paragraph is implemented. Missing: `w:textDirection` on paragraph (CT_TextDirection); RTL-aware default styles; `w:bidi` on section | `word/document.xml` inline | **Fully read** — CT_TextDirection (line 773), ST_TextDirection (tb/rl/lr/tbV/rlV/lrV) | **Batch 5** | Low — `w:bidi` already works; `w:textDirection` is a simple addition |
| 11 | **RTL run direction** | Arabic text runs within mixed-direction paragraphs | Arabic thesis: Arabic text runs; Mixed Arabic/English inline | `rtl` flag exists on `WmlRunProperties` — **already serialized** | `rtl` is implemented. Missing: `cs` (complex script) font size (`szCs`) is implemented. `bCs`/`iCs` implemented. No gaps for basic RTL runs. | `word/document.xml` inline | Already read | **Batch 5** | Low — already mostly implemented |
| 12 | **Arabic-aware font specification** | Correct Arabic font rendering | Arabic thesis: Arabic font for cs/eastAsia slots; Mixed docs | `rFonts.cs` and `rFonts.eastAsia` exist in `WmlFonts` — **already serializable** | No gaps for font specification. Missing: font table (`word/fontTable.xml`) for font embedding declarations | `word/fontTable.xml` (optional but recommended) | CT_FontTable not yet read | **Batch 5** | Low — font table is optional; fonts render without it if installed |
| 13 | **Complex script font size** | Correct Arabic text sizing | Arabic thesis: Arabic text at correct size | `szCs` exists in `WmlRunProperties` — **already serialized** | No gap | Inline | Already read | — | None |
| 14 | **Hyperlinks** | Cross-references, URLs, internal document anchors | Any document with links; Thesis: bibliography URLs | None | `w:hyperlink` not in WmlParagraphContent union; external relationship not wired | Inline in `word/document.xml`; external relationship in `word/_rels/document.xml.rels` | **Fully read** — CT_Hyperlink (line 1218): `r:id`, `anchor`, `tgtFrame`, `tooltip` | **Batch 4** | Low — straightforward; external rels already supported by RelationshipManager |
| 15 | **TOC field** | Automatic table of contents | Thesis: TOC page; Reports: contents page; Any structured doc | None | `w:fldChar`/`w:instrText` not serializable; TOC field instruction string not constructable | Inline in `word/document.xml` | **Fully read** — CT_FldChar (line 1210), CT_SimpleField (line 1158); TOC field is `TOC \o "1-3" \h \z \u` instruction string | **Batch 4** | Medium — TOC field is a complex field sequence; Word must update it on open; `updateFields` in settings needed |
| 16 | **Caption mechanism** | Figure/table/appendix captions for generated lists | Thesis: Figure 1, Table 2 captions; Reports: exhibit captions | None | No caption paragraph style; no SEQ field support; no caption label mechanism | Inline in `word/document.xml`; SEQ field via `w:fldChar`/`w:instrText` | **Fully read** — CT_Caption (line 3407) in settings; SEQ field via CT_FldChar | **Batch 4** | Medium — captions are styled paragraphs + SEQ field; requires field infrastructure |
| 17 | **List of Figures / Tables / Appendices** | Generated lists referencing captioned items | Thesis: list of figures, list of tables; Reports: exhibit index | None | No `TC` field support; no `TOC \c` field for generated lists | Inline in `word/document.xml` | **Fully read** — uses same CT_FldChar/instrText mechanism as TOC | **Batch 4** | Medium — depends on caption + field infrastructure from Batch 4 |
| 18 | **Paragraph textDirection** | Vertical text, explicit LR/RL direction on paragraph | Arabic docs needing explicit paragraph direction beyond `w:bidi` | `bidi` on paragraph works | `w:textDirection` on paragraph not yet serialized (CT_TextDirection exists in CT_PPrBase but not in WmlParagraphProperties) | Inline in `word/document.xml` | **Fully read** — CT_TextDirection (line 773), used in CT_PPrBase (line 1078) | **Batch 5** | Low |
| 19 | **Section textDirection** | RTL section layout | Arabic document sections | `bidi` on section works | `w:textDirection` on section not yet serialized | Inline in `w:sectPr` | **Fully read** — CT_TextDirection in EG_SectPrContents (line 1547) | **Batch 5** | Low |
| 20 | **Even/odd headers** | Different headers on even/odd pages | Thesis: chapter title on odd, blank/logo on even; Books | None | `evenAndOddHeaders` in settings not emitted; even/odd header parts not wired | `word/settings.xml` + `word/header2.xml` (even) | **Fully read** — CT_Settings `evenAndOddHeaders` (line 2784); ST_HdrFtr `even`/`default`/`first` | **Batch 2** | Medium — requires settings part + correct header type wiring |
| 21 | **Footnote/endnote properties** | Control numbering format, restart behavior, position | Thesis: footnotes restart per page; Academic: endnotes at section end | None | `footnotePr`/`endnotePr` in sectPr not serialized; global footnote props in settings not emitted | Inline in `w:sectPr` and `word/settings.xml` | **Fully read** — CT_FtnProps (line 2506), CT_EdnProps (line 2513), CT_FtnDocProps (line 2519) | **Batch 3** | Low |
| 22 | **Table cell borders** | Cell-level border control | Any table with custom cell borders | `WmlTableCellProperties` partial — tcBorders not yet supported | `w:tcBorders` not serialized; CT_TcPr detail not fully read | Inline in `word/document.xml` | **Partially read** — CT_TcPr inferred; tcBorders structure not read | **Batch 1** | Low — read CT_TcPr fully then add |
| 23 | **Table cell margins** | Cell padding control | Any table with custom cell padding | Not supported | `w:tcMar` not serialized | Inline in `word/document.xml` | **Partially read** — CT_TcMar referenced at line 2245 but not read | **Batch 1** | Low |
| 24 | **Document default styles** | Base font, paragraph defaults for entire document | Any document needing consistent defaults; Arabic docs needing Arabic default font | `docDefaults` field in WmlStyles exists but not serialized | `w:docDefaults` with `w:rPrDefault`/`w:pPrDefault` not emitted | Inline in `word/styles.xml` | **Fully read** — CT_DocDefaults (line 2688), CT_RPrDefault (line 2677), CT_PPrDefault (line 2682) | **Batch 5** | Low |
| 25 | **Style docDefaults for Arabic** | Arabic-aware document-level defaults | Arabic thesis: Arabic font as document default; RTL as default paragraph direction | Not supported | `w:docDefaults` not emitted; cannot set Arabic font as document default | Inline in `word/styles.xml` | **Fully read** — same CT_DocDefaults | **Batch 5** | Low |

---

## Summary by Batch

| Batch | Mechanisms | Rationale |
|---|---|---|
| **Batch 1** | Numbering definitions, section breaks + pgNumType, document settings (minimal), table cell borders/margins | Foundation for all structured documents; numbering needed by lists and headings; settings needed by RTL and headers |
| **Batch 2** | Header parts, footer parts, page number fields, even/odd header support | Paginated documents; page numbers in headers/footers |
| **Batch 3** | Footnotes, endnotes, footnote/endnote properties | Academic and legal documents |
| **Batch 4** | Hyperlinks, field infrastructure (fldChar/instrText), TOC field, caption mechanism, generated lists | Advanced document automation |
| **Batch 5** | RTL/Arabic refinements (textDirection, docDefaults, Arabic font defaults), font table | Arabic-first document generation |

---

## What is already working (v1)

- OPC package, content types, relationships
- Paragraphs, runs, text, breaks, tabs
- Full run formatting (bold, italic, underline, color, font, size, highlight, vertAlign, lang, rtl, cs, szCs, bCs, iCs)
- Paragraph formatting (alignment, spacing, indentation, outlineLvl, bidi, keepNext, keepLines)
- Tables (borders, column widths, cell shading, gridSpan, vAlign)
- Styles (paragraph + character, basedOn, next)
- Section properties (pgSz, pgMar, orient, bidi, titlePg, vAlign, type)
- Core and extended document properties
- 113 automated tests passing

## What is NOT in scope for this package (ever)

- Thesis cover page content
- Fixed chapter structure
- University/faculty blocks
- Hardcoded bibliography sections
- Any document-product-specific content

These belong in upstream applications that consume this package.
