# SPEC_TRACEABILITY.md

Code traceability ledger for the OOXML TypeScript library.
Every nontrivial TypeScript construct must have an entry here before implementation.

**Legend:**
- Status: `planned` | `implemented` | `deferred` | `rejected`
- Priority: P0 (must-have v1) | P1 (important v1) | P2 (nice-to-have v1) | P3 (future)

---

## Module: `src/core/` — Shared OOXML Core

### OpcPackage
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART2/opc-contentTypes.xsd, opc-relationships.xsd |
| Reason | Manages the ZIP container, content types registry, and relationship files for any OOXML package |
| Mapping | Wraps JSZip; exposes addPart(), addRelationship(), setContentType(), build() |
| Status | implemented — `src/core/package.ts` |
| Priority | P0 |

### ContentTypeManager
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART2/opc-contentTypes.xsd — CT_Types, CT_Default, CT_Override |
| Reason | Serializes [Content_Types].xml; tracks Default (by extension) and Override (by part URI) entries |
| Mapping | CT_Types → ContentTypeManager; CT_Default → addDefault(ext, contentType); CT_Override → addOverride(partName, contentType) |
| Status | implemented — `src/core/content-types.ts` |
| Priority | P0 |

### RelationshipManager
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART2/opc-relationships.xsd — CT_Relationships, CT_Relationship |
| Reason | Serializes .rels files; manages Id generation and uniqueness (xsd:ID constraint) |
| Mapping | CT_Relationship → { id, type, target, targetMode } |
| Status | implemented — `src/core/relationships.ts` |
| Priority | P0 |

### OpcRelationship (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART2/opc-relationships.xsd — CT_Relationship |
| Reason | Typed representation of a single relationship entry |
| Mapping | Id→id, Type→type, Target→target, TargetMode→targetMode |
| Status | implemented — `src/core/relationships.ts` |
| Priority | P0 |

### OpcCoreProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART2/opc-coreProperties.xsd — CT_CoreProperties |
| Reason | Typed representation of docProps/core.xml content |
| Mapping | dc:title→title, dc:creator→creator, dcterms:created→created, dcterms:modified→modified, etc. |
| Status | implemented — `src/word/builder.ts` as DocxCoreProperties |
| Priority | P1 |

### ExtendedProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/shared-documentPropertiesExtended.xsd — CT_Properties |
| Reason | Typed representation of docProps/app.xml content |
| Mapping | Application, AppVersion, Pages, Words, Characters, Company, etc. |
| Status | implemented — `src/word/builder.ts` as DocxExtendedProperties |
| Priority | P1 |

---

## Module: `src/shared/` — Shared Type Aliases

### XmlBuilder
| Field | Value |
|---|---|
| Kind | class |
| Source | ECMA-376 5th ed. — XML serialization requirement derived from all CT_* types in wml.xsd, opc-*.xsd |
| Reason | All OOXML parts are XML documents. A minimal, dependency-free XML builder is needed to serialize the object model to well-formed XML strings with correct namespace declarations. |
| Mapping | Produces XML strings consumed by OpcPackage/WmlSerializer. Must handle: element creation, attribute setting, text nodes, namespace prefixes, xml:space='preserve' (CT_Text constraint from wml.xsd). |
| Status | implemented — `src/shared/xml.ts` |
| Priority | P0 |

### Twips (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/shared-commonSimpleTypes.xsd — ST_TwipsMeasure |
| Reason | Primary unit for page dimensions, margins, spacing in WML. 1 twip = 1/1440 inch. |
| Mapping | `type Twips = number` — integer twips value |
| Status | implemented — `src/shared/types.ts` |
| Priority | P0 |

### HexColorRGB (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/shared-commonSimpleTypes.xsd — ST_HexColorRGB |
| Reason | 6-char hex color string used in CT_Color, CT_Shd, CT_Border |
| Mapping | `type HexColorRGB = string` — e.g. "FF0000" |
| Status | implemented — `src/shared/types.ts` |
| Priority | P1 |

### RelationshipId (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/shared-relationshipReference.xsd — ST_RelationshipId |
| Reason | Used on r:id, r:embed, r:blip etc. to reference relationships |
| Mapping | `type RelationshipId = string` |
| Status | implemented — `src/shared/types.ts` |
| Priority | P0 |

### ConformanceClass (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/shared-commonSimpleTypes.xsd — ST_ConformanceClass |
| Reason | document/@conformance attribute — strict vs transitional |
| Mapping | `type ConformanceClass = 'strict' \| 'transitional'` |
| Status | implemented — `src/shared/types.ts` |
| Priority | P1 |

### LanguageTag (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/shared-commonSimpleTypes.xsd — ST_Lang |
| Reason | BCP47 language tag used in CT_Language (rPr/lang) |
| Mapping | `type LanguageTag = string` |
| Status | implemented — `src/shared/types.ts` |
| Priority | P1 |

---

## Module: `src/word/` — WordprocessingML

### WmlDocument (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Document (line 3439) |
| Reason | Root of word/document.xml |
| Mapping | CT_Document → { conformance?, body? } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlBody (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Body (line 3233) |
| Reason | Document body containing block-level content |
| Mapping | CT_Body → { content: WmlBlockElement[], sectPr?: WmlSectionProperties } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlBlockElement (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — EG_ContentBlockContent (line 2033) |
| Reason | Union of all block-level elements: paragraph, table |
| Mapping | `type WmlBlockElement = WmlParagraph \| WmlTable` (v1 scope; sdt/customXml deferred) |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlParagraph (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_P (line 2159) |
| Reason | Core block element — w:p |
| Mapping | CT_P → { pPr?: WmlParagraphProperties, content: WmlParagraphContent[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlParagraphProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_PPr (line 1038), CT_PPrBase (line 1049) |
| Reason | Paragraph formatting — w:pPr |
| Mapping | pStyle, keepNext, keepLines, pageBreakBefore, numPr, pBdr, shd, tabs, spacing, ind, jc, outlineLvl, sectPr |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlRun (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_R (line 1703) |
| Reason | Inline text run — w:r |
| Mapping | CT_R → { rPr?: WmlRunProperties, content: WmlRunContent[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlRunProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_RPr (line 1790), EG_RPrBase (line 1742) |
| Reason | Character formatting — w:rPr |
| Mapping | rStyle, rFonts, b, i, caps, smallCaps, strike, color, sz, szCs, highlight, u, vertAlign, lang, spacing, kern, position, shd, bdr, vanish, rtl |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlText (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Text (line ~1660) |
| Reason | Text content of a run — w:t |
| Mapping | CT_Text → { text: string, space?: 'preserve' \| 'default' } |
| Note | xml:space='preserve' MUST be set when text has leading/trailing whitespace |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlBreak (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Br (line ~1600) |
| Reason | Break element — w:br |
| Mapping | CT_Br → { type?: 'page' \| 'column' \| 'textWrapping', clear?: 'none' \| 'left' \| 'right' \| 'all' } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlTable (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Tbl (line 2434) |
| Reason | Table element — w:tbl |
| Mapping | CT_Tbl → { tblPr: WmlTableProperties, tblGrid: WmlTableGrid, rows: WmlTableRow[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlTableRow (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Row (line ~2330) |
| Reason | Table row — w:tr |
| Mapping | CT_Row → { trPr?: WmlTableRowProperties, cells: WmlTableCell[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlTableCell (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Tc (line 2270) |
| Reason | Table cell — w:tc. MUST contain at least one block element. |
| Mapping | CT_Tc → { tcPr?: WmlTableCellProperties, content: WmlBlockElement[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlTableProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_TblPrBase (line ~2390) |
| Reason | Table formatting — w:tblPr |
| Mapping | tblStyle, tblW, jc, tblBorders, shd, tblLayout, tblCellMar, tblLook |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlTableGrid (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_TblGrid |
| Reason | Column width definitions — w:tblGrid |
| Mapping | CT_TblGrid → { cols: Array<{ w: Twips }> } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlSectionProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_SectPr (line 1566), EG_SectPrContents |
| Reason | Section layout — w:sectPr. Controls page size, margins, headers/footers. |
| Mapping | type, pgSz, pgMar, cols, titlePg, pgNumType, headerReference[], footerReference[] |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlPageSize (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_PageSz (line 1367) |
| Reason | Page dimensions — w:pgSz |
| Mapping | CT_PageSz → { w?: Twips, h?: Twips, orient?: 'portrait' \| 'landscape' } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlPageMargins (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_PageMar (line 1373) |
| Reason | Page margins — w:pgMar |
| Mapping | CT_PageMar → { top, right, bottom, left, header, footer, gutter } all Twips |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlStyles (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Styles (line ~3110) |
| Reason | Root of word/styles.xml |
| Mapping | CT_Styles → { docDefaults?, styles: WmlStyle[] } |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlStyle (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Style (line 3058) |
| Reason | Style definition — w:style |
| Mapping | type, styleId, default, name, basedOn, next, link, pPr, rPr, tblPr, trPr, tcPr |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlStyleType (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_StyleType |
| Reason | Discriminates paragraph/character/table/numbering styles |
| Mapping | `type WmlStyleType = 'paragraph' \| 'character' \| 'table' \| 'numbering'` |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

### WmlJustificationValue (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_Jc |
| Reason | Paragraph alignment values |
| Mapping | `'start' \| 'center' \| 'end' \| 'both' \| 'distribute' \| ...` |
| Status | implemented — `src/word/types.ts` |
| Priority | P1 |

### WmlSerializer
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART1/wml.xsd — all CT_* types |
| Reason | Converts WML object model to XML strings for inclusion in DOCX ZIP |
| Mapping | serializeDocument(), serializeStyles(), serializeNumbering(), serializeSettings() |
| Status | implemented — `src/word/serializer.ts` |
| Priority | P0 |

### DocxBuilder
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART1/wml.xsd + PART2/opc-*.xsd |
| Reason | High-level API: addParagraph(), addTable(), setPageSize(), setMargins(), addStyle(), build() → Buffer |
| Mapping | Orchestrates WmlSerializer + OpcPackage to produce valid DOCX |
| Status | implemented — `src/word/builder.ts` |
| Priority | P0 |

---

## Deferred Constructs

| Construct | Source file (per manifest) | Reason | Future Module | Target version |
|---|---|---|---|---|
| WmlNumbering, WmlAbstractNum, WmlLevel | 5TH EDITION PART1/wml.xsd (CT_Numbering, CT_AbstractNum, CT_Lvl — **now read** lines 2955-3030) | List/numbering support | src/word/ | **Batch 1** |
| WmlFootnote, WmlEndnote | 5TH EDITION PART1/wml.xsd (CT_FtnEdn — read lines 2480-2560) | Footnote/endnote support | src/word/ | **Batch 3 — IMPLEMENTED** |
| WmlDrawingInline, WmlDrawingAnchor | 5TH EDITION PART1/dml-wordprocessingDrawing.xsd (unread), dml-main.xsd (unread), dml-picture.xsd (unread) | Inline image embedding | src/word/drawing/ | v2 |
| WmlFldChar, WmlInstrText | 5TH EDITION PART1/wml.xsd (CT_FldChar — **now read** lines 1155-1260) | Field infrastructure | src/word/ | **Batch 2** |
| WmlHyperlink | 5TH EDITION PART1/wml.xsd (CT_Hyperlink — **now read** line 1218) | Hyperlink support | src/word/ | **Batch 4** |
| WmlSettings / DocxSettings | 5TH EDITION PART1/wml.xsd (CT_Settings — **now read** lines 2733-2840) | Document settings part | src/word/ | **Batch 1** |
| WmlHeaderFooterContent | 5TH EDITION PART1/wml.xsd (CT_HdrFtr — **now read** line 1528) | Header/footer parts | src/word/ | **Batch 2** |
| WmlPageNumbering | 5TH EDITION PART1/wml.xsd (CT_PageNumber — **now read** line 1457) | Page number format/start | src/word/ | **Batch 1** |
| WmlTextDirectionValue | 5TH EDITION PART1/wml.xsd (CT_TextDirection — **now read** line 773) | RTL/Arabic text direction | src/word/ | **Batch 5** |
| WmlDocumentDefaults | 5TH EDITION PART1/wml.xsd (CT_DocDefaults — read line 2688) | Document-level style defaults | src/word/ | **Batch 5** |
| WmlFontTable | 5TH EDITION PART1/wml.xsd (CT_FontTable — not in XSD, only in prose spec) | Font table part | src/word/ | **Batch 5 (deferred - insufficient schema-grounded source support)** |
| WmlComments | 5TH EDITION PART1/wml.xsd (CT_Comments — not yet read) | Comments part | src/word/ | v2 |
| SmlWorkbook, SmlWorksheet, SmlCell | 5TH EDITION PART1/sml.xsd (classified, deferred) | SpreadsheetML / XLSX | src/xlsx/ | v2 |
| PmlPresentation, PmlSlide | 5TH EDITION PART1/pml.xsd (classified, deferred) | PresentationML / PPTX | src/pptx/ | v3 |
| OpcDigitalSignature | 5TH EDITION PART2/opc-digSig.xsd (classified, deferred) | Package signing | src/core/ | v2 |
| WmlBibliography | 5TH EDITION PART1/shared-bibliography.xsd (classified, deferred) | Citations | src/word/ | v2 |

---

## Planned Constructs — Next Implementation Batches

### Batch 1: Numbering, Section Mechanics, Settings, Table Cell Completion

#### WmlNumberFormat (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_NumberFormat (line 1294) |
| Reason | Used by numbering levels, page numbering, footnote numbering |
| Mapping | `'decimal' \| 'upperRoman' \| 'lowerRoman' \| 'bullet' \| 'none' \| 'arabicAbjad' \| 'arabicAlpha' \| ...` |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### WmlNumberingLevel (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Lvl (line 2961) |
| Reason | Defines one level of a numbering definition (format, text pattern, indent, style) |
| Mapping | ilvl, start, numFmt, lvlText, lvlJc, pPr, rPr, pStyle, suff |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### WmlAbstractNum (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_AbstractNum (line 2989) |
| Reason | Abstract numbering definition — reusable template for list formats |
| Mapping | abstractNumId, multiLevelType, levels (CT_Lvl[]) |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### WmlNum (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Num (line 3004) |
| Reason | Concrete numbering instance — references abstractNum, used by paragraphs via numId |
| Mapping | numId, abstractNumId, lvlOverride[] |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### WmlNumbering (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Numbering (line 3015) |
| Reason | Root of word/numbering.xml |
| Mapping | abstractNums[], nums[] |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### serializeNumbering (function)
| Field | Value |
|---|---|
| Kind | function |
| Source | 5TH EDITION PART1/wml.xsd — CT_Numbering |
| Reason | Produces word/numbering.xml string |
| Mapping | WmlNumbering → XML string |
| Status | implemented — `src/word/numbering.ts` |
| Priority | P0 |

#### WmlPageNumbering (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_PageNumber (line 1457) |
| Reason | Controls page number format and start value per section (w:pgNumType) |
| Mapping | fmt (WmlNumberFormat), start (number) |
| Status | implemented — `src/word/types.ts` |
| Priority | P0 |

#### DocxSettings (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Settings (line 2733) |
| Reason | Controls document-level behavior: field update, even/odd headers, footnote config, RTL |
| Mapping | evenAndOddHeaders, updateFields, footnotePr, endnotePr, themeFontLang, defaultTabStop |
| Status | implemented — `src/word/settings.ts` |
| Priority | P0 |

#### serializeSettings (function)
| Field | Value |
|---|---|
| Kind | function |
| Source | 5TH EDITION PART1/wml.xsd — CT_Settings |
| Reason | Produces word/settings.xml string |
| Status | implemented — `src/word/settings.ts` |
| Priority | P0 |

### Batch 2: Headers, Footers, Page Number Fields

#### WmlHeaderFooterType (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_HdrFtr |
| Reason | Discriminates default/first/even header-footer variants |
| Mapping | `'default' \| 'first' \| 'even'` |
| Status | implemented — `src/word/header-footer.ts` |
| Priority | P0 |

#### WmlHeaderFooterRef (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_HdrFtrRef (line 1514) |
| Reason | Links a section to a header/footer part via relationship ID and type |
| Mapping | type (WmlHeaderFooterType), id (RelationshipId) |
| Status | implemented — `src/word/header-footer.ts` |
| Priority | P0 |

#### serializeHeader / serializeFooter (functions)
| Field | Value |
|---|---|
| Kind | function |
| Source | 5TH EDITION PART1/wml.xsd — CT_HdrFtr (line 1528) |
| Reason | Produces word/header*.xml and word/footer*.xml strings |
| Status | implemented — `src/word/header-footer.ts` |
| Priority | P0 |

#### WmlFldChar (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_FldChar (line 1210) |
| Reason | Field character run content — begin/separate/end markers for complex fields |
| Mapping | _type: 'fldChar', fldCharType: 'begin'\|'separate'\|'end', fldLock?, dirty? |
| Status | implemented — `src/word/fields.ts` |
| Priority | P0 |

#### WmlInstrText (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — instrText element in EG_RunInnerContent |
| Reason | Field instruction text run content — carries the field instruction string |
| Mapping | _type: 'instrText', text: string, space?: 'preserve' |
| Status | implemented — `src/word/fields.ts` |
| Priority | P0 |

#### pageNumberField / totalPagesField / fieldRuns (functions)
| Field | Value |
|---|---|
| Kind | factory functions |
| Source | 5TH EDITION PART1/wml.xsd — CT_FldChar, instrText |
| Reason | Reusable helpers for PAGE, NUMPAGES, and generic field sequences |
| Status | implemented — `src/word/fields.ts` |
| Priority | P0 |

### Batch 3: Footnotes and Endnotes

#### WmlNoteType (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_FtnEdn (~line 2470) |
| Reason | Discriminates normal user notes from required system entries (separator, continuationSeparator) |
| Mapping | `'normal' \| 'separator' \| 'continuationSeparator' \| 'continuationNotice'` |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P0 |

#### WmlFootnote (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnEdn (line 2491) |
| Reason | Individual footnote/endnote entry in footnotes.xml / endnotes.xml |
| Mapping | id, type (normal/separator/continuationSeparator), content (WmlBlockElement[]) |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P0 |

#### WmlFootnoteRef (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnEdnRef (line 2484) |
| Reason | Run content element that places a footnote/endnote reference mark in the body |
| Mapping | _type: 'footnoteReference'\|'endnoteReference', id: number |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P0 |

#### WmlFootnoteProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnProps (line 2506) |
| Reason | Controls footnote position, number format, and restart behavior |
| Mapping | pos, numFmt, numStart, numRestart |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P1 |

#### WmlEndnoteProperties (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_EdnProps (line 2513) |
| Reason | Controls endnote position, number format, and restart behavior |
| Mapping | pos, numFmt, numStart, numRestart |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P1 |

#### NoteManager (class)
| Field | Value |
|---|---|
| Kind | class |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnEdn, CT_FtnEdnRef |
| Reason | Tracks footnotes and endnotes, assigns sequential IDs, returns reference run content |
| Mapping | addFootnote() → WmlFootnoteRef (id=1+), addEndnote() → WmlFootnoteRef (id=1+, independent) |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P0 |

#### serializeFootnotes / serializeEndnotes (functions)
| Field | Value |
|---|---|
| Kind | functions |
| Source | 5TH EDITION PART1/wml.xsd — CT_Footnotes (line 3245), CT_Endnotes (line 3253) |
| Reason | Produces word/footnotes.xml and word/endnotes.xml; always includes required separator entries |
| Status | **implemented** — `src/word/footnotes.ts` |
| Priority | P0 |

#### DocxBuilder.addFootnote / addEndnote (methods)
| Field | Value |
|---|---|
| Kind | methods |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnEdn, CT_FtnEdnRef |
| Reason | High-level API: add note content, get back a WmlRun with reference mark to insert in body |
| Mapping | addFootnote(content) → WmlRun { _type:'run', content:[{_type:'footnoteReference', id}] } |
| Status | **implemented** — `src/word/builder.ts` |
| Priority | P0 |

#### DocxBuilder.configureFootnotes / configureEndnotes (methods)
| Field | Value |
|---|---|
| Kind | methods |
| Source | 5TH EDITION PART1/wml.xsd — CT_FtnProps (line 2506), CT_EdnProps (line 2513) |
| Reason | Configure note position, number format, restart behavior |
| Status | **implemented** — `src/word/builder.ts` |
| Priority | P1 |

### Batch 4: Hyperlinks, TOC, Captions

#### WmlHyperlink (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_Hyperlink (line 1218) |
| Reason | Inline hyperlink in paragraph content — external URL or internal anchor |
| Mapping | _type: 'hyperlink', id? (r:id for external), anchor? (internal), content: WmlRun[] |
| Status | implemented — src/word/hyperlinks.ts |
| Priority | P0 |

#### tableOfContents / captionParagraph / listOfFigures / listOfTables (functions)
| Field | Value |
|---|---|
| Kind | factory functions |
| Source | 5TH EDITION PART1/wml.xsd — CT_FldChar, instrText (TOC/SEQ fields) |
| Reason | Reusable document automation helpers for structured documents |
| Status | implemented — `src/word/helpers.ts`, `src/word/fields.ts` |
| Priority | P1 |

### Batch 5: RTL/Arabic and Document Defaults

#### WmlTextDirectionValue (type alias)
| Field | Value |
|---|---|
| Kind | type alias |
| Source | 5TH EDITION PART1/wml.xsd — ST_TextDirection (line 765) |
| Reason | Explicit text direction for paragraphs, table cells, sections |
| Mapping | `'tb' \| 'rl' \| 'lr' \| 'tbV' \| 'rlV' \| 'lrV'` |
| Status | implemented — `src/word/types.ts` |
| Priority | P1 |

#### WmlDocumentDefaults (interface)
| Field | Value |
|---|---|
| Kind | interface |
| Source | 5TH EDITION PART1/wml.xsd — CT_DocDefaults (line 2688) |
| Reason | Document-level default run and paragraph properties in styles.xml |
| Mapping | rPrDefault (WmlRunProperties), pPrDefault (WmlParagraphProperties) |
| Status | implemented — `src/word/types.ts` |
| Priority | P1 |

#### serializeDocDefaults (function)
| Field | Value |
|---|---|
| Kind | function |
| Source | 5TH EDITION PART1/wml.xsd — CT_DocDefaults serialization |
| Reason | Serialize WmlDocDefaults to w:docDefaults XML element |
| Mapping | WmlDocDefaults → `<w:docDefaults><w:rPrDefault>...</w:rPrDefault><w:pPrDefault>...</w:pPrDefault></w:docDefaults>` |
| Status | implemented — `src/word/serializer.ts` |
| Priority | P1 |

#### DocxBuilder.configureDocDefaults (method)
| Field | Value |
|---|---|
| Kind | method |
| Source | 5TH EDITION PART1/wml.xsd — CT_DocDefaults integration |
| Reason | Configure document-level defaults on DocxBuilder |
| Mapping | WmlDocDefaults → builder.docDefaults field → styles.xml w:docDefaults |
| Status | implemented — `src/word/builder.ts` |
| Priority | P1 |

#### DocxBuilder.setSectionTextDirection (method)
| Field | Value |
|---|---|
| Kind | method |
| Source | 5TH EDITION PART1/wml.xsd — CT_TextDirection on sectPr |
| Reason | Set text direction for document sections |
| Mapping | WmlTextDirectionValue → sectPr.textDirection → `<w:textDirection w:val="..."/>` |
| Status | implemented — `src/word/builder.ts` |
| Priority | P1 |
