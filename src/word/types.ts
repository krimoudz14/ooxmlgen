/**
 * WordprocessingML v1 type definitions.
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd (partially read — see repo-manifest.json)
 * Traceability: SPEC_TRACEABILITY.md — WmlDocument through WmlStyle
 *
 * V1 BOUNDARY: This file covers only the constructs whose source sections
 * in wml.xsd were actually read. See SPEC_TRACEABILITY.md deferred table
 * for what is explicitly excluded.
 *
 * DEFERRED (not in this file):
 *   - WmlNumbering / WmlAbstractNum / WmlLevel  → v1.1 (CT_Numbering unread)
 *   - WmlFootnote / WmlEndnote                  → v1.1 (CT_FtnEdn unread)
 *   - WmlDrawing / WmlInline / WmlAnchor        → v1.1 (dml-wordprocessingDrawing.xsd unread)
 *   - WmlField / WmlSimpleField                 → v1.1 (CT_FldChar unread)
 *   - WmlSettings                               → v1.1 (CT_Settings unread)
 *   - WmlFontTable                              → v1.1 (CT_FontTable unread)
 *   - WmlComments                               → v1.1 (CT_Comments unread)
 */

import type { Twips, HalfPoints, HexColor, SignedTwips, LanguageTag, ConformanceClass } from '../shared/types';

// ---------------------------------------------------------------------------
// Batch 1 additions — new primitive types
// ---------------------------------------------------------------------------

/**
 * Source: ST_TextDirection (wml.xsd line 765)
 * Explicit text direction for paragraphs, table cells, and sections.
 * Key values: 'lr' = left-to-right, 'rl' = right-to-left, 'tb' = top-to-bottom.
 */
export type WmlTextDirectionValue = 'tb' | 'rl' | 'lr' | 'tbV' | 'rlV' | 'lrV';

/**
 * Source: CT_PageNumber (wml.xsd line 1457)
 * Page number format and start value for a section.
 * Used in w:sectPr/w:pgNumType.
 */
export interface WmlPageNumbering {
  /** Number format. Default 'decimal'. arabicAbjad for Arabic documents. */
  fmt?: import('./numbering').WmlNumberFormat;
  /** Starting page number for this section. */
  start?: number;
}

// ---------------------------------------------------------------------------
// Shared sub-types (sourced from wml.xsd read sections)
// ---------------------------------------------------------------------------

/**
 * Source: CT_Color (wml.xsd ~line 150)
 * val: HexColor | 'auto'; themeColor/themeTint/themeShade optional (deferred)
 */
export interface WmlColor {
  val: HexColor;
}

/**
 * Source: CT_Fonts (wml.xsd ~line 1730)
 * Font family specification for a run.
 */
export interface WmlFonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
}

/**
 * Source: CT_Spacing (wml.xsd ~line 620)
 * Paragraph spacing before/after and line spacing.
 * All values in twips. lineRule: 'auto' | 'exact' | 'atLeast'
 */
export interface WmlSpacing {
  before?: Twips;
  after?: Twips;
  line?: SignedTwips;
  lineRule?: 'auto' | 'exact' | 'atLeast';
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
}

/**
 * Source: CT_Ind (wml.xsd ~line 635)
 * Paragraph indentation. All values in twips.
 */
export interface WmlIndentation {
  start?: SignedTwips;
  end?: SignedTwips;
  hanging?: Twips;
  firstLine?: Twips;
}

/**
 * Source: ST_Jc (wml.xsd ~line 650), CT_Jc
 * Paragraph alignment.
 */
export type WmlJustificationValue = 'start' | 'center' | 'end' | 'both' | 'distribute' | 'numTab' | 'highKashida' | 'lowKashida' | 'mediumKashida' | 'thaiDistribute';

/**
 * Source: CT_Border (wml.xsd ~line 415)
 * Border specification. val is ST_Border enum value.
 * Art border values (apples, bats, etc.) are valid per spec but decorative.
 */
export interface WmlBorder {
  val: string;           // ST_Border enum — 'single', 'double', 'none', 'nil', etc.
  color?: HexColor;
  sz?: number;           // ST_EighthPointMeasure — border width in 1/8 pt
  space?: number;        // ST_PointMeasure — space from text in pt
  shadow?: boolean;
}

/**
 * Source: CT_Shd (wml.xsd ~line 430)
 * Shading/fill for paragraphs, runs, table cells.
 */
export interface WmlShading {
  val: string;           // ST_Shd enum — 'clear', 'solid', 'nil', etc.
  color?: HexColor;
  fill?: HexColor;
}

/**
 * Numbering properties on a paragraph — points to word/numbering.xml.
 *
 * @provisional DEFERRED_V1_1 — This type exists for forward compatibility.
 * The serializer does NOT emit w:numPr until word/numbering.xml support is
 * added in v1.1. Setting numPr on a paragraph has no effect in v1.
 * Source: CT_NumPr (wml.xsd — inferred from CT_PPrBase child, unread detail)
 */
export interface WmlNumberingProperties {
  ilvl: number;   // CT_DecimalNumber — list level 0-8
  numId: number;  // CT_DecimalNumber — references w:num/@w:numId in numbering.xml
}

// ---------------------------------------------------------------------------
// Run properties
// Source: CT_RPr (wml.xsd line 1790), EG_RPrBase (wml.xsd line 1742)
// ---------------------------------------------------------------------------

/**
 * Source: CT_RPr / EG_RPrBase (wml.xsd lines 1742-1790)
 * Character formatting applied to a run.
 */
export interface WmlRunProperties {
  rStyle?: string;          // CT_String — character style ID
  rFonts?: WmlFonts;        // CT_Fonts
  b?: boolean;              // CT_OnOff — bold
  bCs?: boolean;            // CT_OnOff — bold complex script
  i?: boolean;              // CT_OnOff — italic
  iCs?: boolean;            // CT_OnOff — italic complex script
  caps?: boolean;           // CT_OnOff — all caps
  smallCaps?: boolean;      // CT_OnOff — small caps
  strike?: boolean;         // CT_OnOff — strikethrough
  dstrike?: boolean;        // CT_OnOff — double strikethrough
  outline?: boolean;        // CT_OnOff
  shadow?: boolean;         // CT_OnOff
  emboss?: boolean;         // CT_OnOff
  imprint?: boolean;        // CT_OnOff
  vanish?: boolean;         // CT_OnOff — hidden text
  color?: WmlColor;         // CT_Color
  sz?: HalfPoints;          // CT_HpsMeasure — font size in half-points
  szCs?: HalfPoints;        // CT_HpsMeasure — complex script font size
  highlight?: string;       // CT_Highlight — ST_HighlightColor enum value
  u?: string;               // CT_Underline — ST_Underline enum value ('single', 'double', etc.)
  vertAlign?: 'baseline' | 'superscript' | 'subscript';  // CT_VerticalAlignRun
  lang?: { val?: LanguageTag; eastAsia?: LanguageTag; bidi?: LanguageTag };  // CT_Language
  spacing?: SignedTwips;    // CT_SignedTwipsMeasure — character spacing
  kern?: HalfPoints;        // CT_HpsMeasure — kerning threshold
  position?: number;        // CT_SignedHpsMeasure — vertical position
  shd?: WmlShading;         // CT_Shd
  bdr?: WmlBorder;          // CT_Border — run border
  rtl?: boolean;            // CT_OnOff — right-to-left
  cs?: boolean;             // CT_OnOff — complex script
  noProof?: boolean;        // CT_OnOff — no spell/grammar check
  snapToGrid?: boolean;     // CT_OnOff
  webHidden?: boolean;      // CT_OnOff
}

// ---------------------------------------------------------------------------
// Run content
// Source: EG_RunInnerContent (wml.xsd lines 1667-1700)
// ---------------------------------------------------------------------------

/**
 * Source: CT_Text (wml.xsd ~line 1660)
 * Text content of a run — w:t element.
 * CRITICAL: xml:space="preserve" must be set when text has leading/trailing whitespace.
 */
export interface WmlText {
  _type: 'text';
  text: string;
  /** xml:space attribute. Default 'preserve' — always safe to set. */
  space?: 'preserve' | 'default';
}

/**
 * Source: CT_Br (wml.xsd ~line 1600)
 * Break element — w:br.
 */
export interface WmlBreak {
  _type: 'break';
  breakType?: 'page' | 'column' | 'textWrapping';
  clear?: 'none' | 'left' | 'right' | 'all';
}

/**
 * Source: CT_Empty (wml.xsd line 22) used for w:tab, w:cr
 */
export interface WmlTab {
  _type: 'tab';
}

export interface WmlCarriageReturn {
  _type: 'cr';
}

/**
 * Union of all run content types.
 * Source: EG_RunInnerContent (wml.xsd lines 1667-1700)
 *
 * Batch 2 additions: WmlFldChar, WmlInstrText (for page number fields)
 * Batch 3 additions: WmlFootnoteRef (footnoteReference / endnoteReference)
 *
 * DEFERRED run content (not in this union):
 *   - w:drawing  → DEFERRED_BATCH4+ (dml-wordprocessingDrawing.xsd unread)
 *   - w:object   → OUT_OF_SCOPE
 *   - w:ruby     → DEFERRED_V2
 */
export type WmlRunContent =
  | WmlText
  | WmlBreak
  | WmlTab
  | WmlCarriageReturn
  | import('./fields').WmlFldChar
  | import('./fields').WmlInstrText
  | import('./footnotes').WmlFootnoteRef;

// ---------------------------------------------------------------------------
// Run
// Source: CT_R (wml.xsd line 1703)
// ---------------------------------------------------------------------------

/**
 * Source: CT_R (wml.xsd line 1703)
 * Inline text run — w:r.
 */
export interface WmlRun {
  _type: 'run';
  rPr?: WmlRunProperties;
  content: WmlRunContent[];
}

// ---------------------------------------------------------------------------
// Paragraph content
// Source: EG_PContent (wml.xsd ~line 2150)
// ---------------------------------------------------------------------------

/**
 * Union of all v1 paragraph inline content.
 * Source: EG_PContent (wml.xsd ~line 2150)
 *
 * DEFERRED paragraph content (not in this union):
 *   - w:hyperlink  → DEFERRED_V1_1 (CT_Hyperlink detail unread)
 *   - w:fldSimple  → DEFERRED_V1_1 (CT_SimpleField unread)
 *   - w:sdt        → DEFERRED_V2
 *   - w:customXml  → DEFERRED_V2
 *   - w:subDoc     → OUT_OF_SCOPE
 */
export type WmlParagraphContent =
  | WmlRun
  | import('./hyperlinks').WmlHyperlink;

// ---------------------------------------------------------------------------
// Paragraph properties
// Source: CT_PPrBase (wml.xsd line 1049), CT_PPr (wml.xsd line 1038)
// ---------------------------------------------------------------------------

/**
 * Source: CT_PPrBase (wml.xsd line 1049)
 * Paragraph formatting — w:pPr.
 */
export interface WmlParagraphProperties {
  pStyle?: string;              // CT_String — paragraph style ID
  keepNext?: boolean;           // CT_OnOff
  keepLines?: boolean;          // CT_OnOff
  pageBreakBefore?: boolean;    // CT_OnOff
  widowControl?: boolean;       // CT_OnOff
  suppressLineNumbers?: boolean;// CT_OnOff
  spacing?: WmlSpacing;         // CT_Spacing
  ind?: WmlIndentation;         // CT_Ind
  jc?: WmlJustificationValue;   // CT_Jc
  outlineLvl?: number;          // CT_DecimalNumber — 0-8 (heading levels)
  shd?: WmlShading;             // CT_Shd
  bidi?: boolean;               // CT_OnOff — right-to-left paragraph
  contextualSpacing?: boolean;  // CT_OnOff
  suppressAutoHyphens?: boolean;// CT_OnOff
  /**
   * Numbering properties — now serialized (Batch 1).
   * Source: CT_NumPr (wml.xsd — inferred from CT_PPrBase child)
   */
  numPr?: WmlNumberingProperties;
  /**
   * Explicit text direction for this paragraph.
   * Source: CT_TextDirection (wml.xsd line 773) — Batch 1 addition.
   */
  textDirection?: WmlTextDirectionValue;
  /**
   * Mid-document section break — places a w:sectPr inside this paragraph's pPr.
   * Source: CT_PPr (wml.xsd line 1038) — sectPr child of pPr.
   * When set, this paragraph ends the current section and starts a new one.
   * The body's final sectPr defines the last section.
   * DEFERRED_BATCH2: headerReference/footerReference on this sectPr not yet wired.
   */
  sectPr?: WmlSectionProperties;
}

// ---------------------------------------------------------------------------
// Paragraph
// Source: CT_P (wml.xsd line 2159)
// ---------------------------------------------------------------------------

/**
 * Source: CT_P (wml.xsd line 2159)
 * Block-level paragraph element — w:p.
 */
export interface WmlParagraph {
  _type: 'paragraph';
  pPr?: WmlParagraphProperties;
  content: WmlParagraphContent[];
}

// ---------------------------------------------------------------------------
// Table sub-types
// Source: CT_TblPrBase (~line 2390), CT_TblGrid, CT_TrPrBase, CT_TcPr
// ---------------------------------------------------------------------------

/**
 * Source: CT_TblBorders (wml.xsd ~line 2370)
 */
export interface WmlTableBorders {
  top?: WmlBorder;
  start?: WmlBorder;
  bottom?: WmlBorder;
  end?: WmlBorder;
  insideH?: WmlBorder;
  insideV?: WmlBorder;
}

/**
 * Source: CT_TcBorders (wml.xsd line 2206) — Batch 1 addition
 * Cell-level borders. Includes diagonal borders (tl2br, tr2bl).
 */
export interface WmlTableCellBorders {
  top?: WmlBorder;
  start?: WmlBorder;
  bottom?: WmlBorder;
  end?: WmlBorder;
  insideH?: WmlBorder;
  insideV?: WmlBorder;
  tl2br?: WmlBorder;  // top-left to bottom-right diagonal
  tr2bl?: WmlBorder;  // top-right to bottom-left diagonal
}

/**
 * Source: CT_TcMar (wml.xsd line 2218) — Batch 1 addition
 * Cell margins (padding). Uses CT_TblWidth (w + type).
 */
export interface WmlTableCellMargins {
  top?: WmlTableWidth;
  start?: WmlTableWidth;
  bottom?: WmlTableWidth;
  end?: WmlTableWidth;
}

/**
 * Source: CT_TblWidth (wml.xsd ~line 2180)
 * Table/cell width specification.
 */
export interface WmlTableWidth {
  w: number;                          // width value
  type: 'nil' | 'pct' | 'dxa' | 'auto';  // ST_TblWidth
}

/**
 * Source: CT_TblPrBase (wml.xsd ~line 2390)
 * Table properties — w:tblPr.
 */
export interface WmlTableProperties {
  tblStyle?: string;              // CT_String — table style ID
  tblW?: WmlTableWidth;           // CT_TblWidth
  jc?: 'start' | 'center' | 'end';  // CT_JcTable — ST_JcTable
  tblBorders?: WmlTableBorders;   // CT_TblBorders
  shd?: WmlShading;               // CT_Shd
  tblLayout?: 'fixed' | 'autofit'; // CT_TblLayoutType
}

/**
 * Source: CT_TblGrid (wml.xsd ~line 2210), CT_TblGridCol
 * Column width definitions — w:tblGrid.
 */
export interface WmlTableGrid {
  cols: Array<{ w: Twips }>;  // CT_TblGridCol — each column width in twips
}

/**
 * Source: CT_TrPrBase (wml.xsd ~line 2300)
 * Table row properties — w:trPr.
 */
export interface WmlTableRowProperties {
  cantSplit?: boolean;          // CT_OnOff — row cannot split across pages
  tblHeader?: boolean;          // CT_OnOff — repeat as header row
  trHeight?: { val: Twips; hRule?: 'auto' | 'exact' | 'atLeast' };  // CT_Height
  jc?: 'start' | 'center' | 'end';  // CT_JcTable
}

/**
 * Table cell properties — w:tcPr.
 *
 * Batch 1: Now fully implemented.
 * Source: CT_TcPrBase (wml.xsd line 2236) — all children now read and supported.
 */
export interface WmlTableCellProperties {
  tcW?: WmlTableWidth;          // CT_TblWidth — cell width
  gridSpan?: number;            // CT_DecimalNumber — column span
  /** Vertical merge. 'restart' = start of merge, 'continue' = continuation cell. */
  vMerge?: 'restart' | 'continue';  // CT_VMerge
  tcBorders?: WmlTableCellBorders;  // CT_TcBorders — Batch 1 addition
  shd?: WmlShading;             // CT_Shd — cell shading
  noWrap?: boolean;             // CT_OnOff — prevent text wrapping
  tcMar?: WmlTableCellMargins;  // CT_TcMar — Batch 1 addition
  /** Explicit text direction for this cell. */
  textDirection?: WmlTextDirectionValue;  // CT_TextDirection — Batch 1 addition
  vAlign?: 'top' | 'center' | 'bottom';  // CT_VerticalJc
  hideMark?: boolean;           // CT_OnOff — hide end-of-cell mark
}

// ---------------------------------------------------------------------------
// Table
// Source: CT_Tbl (wml.xsd line 2434), CT_Row, CT_Tc
// ---------------------------------------------------------------------------

/**
 * Source: CT_Tc (wml.xsd line 2270)
 * Table cell — w:tc.
 * SPEC CONSTRAINT: must contain at least one block-level element.
 */
export interface WmlTableCell {
  _type: 'tableCell';
  tcPr?: WmlTableCellProperties;
  content: WmlBlockElement[];   // EG_BlockLevelElts — min 1 required
}

/**
 * Source: CT_Row (wml.xsd ~line 2330)
 * Table row — w:tr.
 */
export interface WmlTableRow {
  _type: 'tableRow';
  trPr?: WmlTableRowProperties;
  cells: WmlTableCell[];
}

/**
 * Source: CT_Tbl (wml.xsd line 2434)
 * Table element — w:tbl.
 * SPEC CONSTRAINT: tblPr and tblGrid are required children.
 */
export interface WmlTable {
  _type: 'table';
  tblPr: WmlTableProperties;
  tblGrid: WmlTableGrid;
  rows: WmlTableRow[];
}

// ---------------------------------------------------------------------------
// Block-level content union
// Source: EG_ContentBlockContent (wml.xsd line 2033)
// ---------------------------------------------------------------------------

/**
 * Union of all v1 block-level elements.
 * Source: EG_ContentBlockContent (wml.xsd line 2033)
 *
 * DEFERRED block content (not in this union):
 *   - w:sdt        → DEFERRED_V2
 *   - w:customXml  → DEFERRED_V2
 *   - w:altChunk   → OUT_OF_SCOPE
 */
export type WmlBlockElement = WmlParagraph | WmlTable;

// ---------------------------------------------------------------------------
// Section properties
// Source: CT_SectPr (wml.xsd line 1566), EG_SectPrContents
// ---------------------------------------------------------------------------

/**
 * Source: CT_PageSz (wml.xsd line 1367)
 * Page dimensions — w:pgSz.
 * Common values: A4 = {w:11906, h:16838}, US Letter = {w:12240, h:15840}
 */
export interface WmlPageSize {
  w?: Twips;
  h?: Twips;
  orient?: 'portrait' | 'landscape';
}

/**
 * Source: CT_PageMar (wml.xsd line 1373)
 * Page margins — w:pgMar. All values in twips. All required per schema.
 */
export interface WmlPageMargins {
  top: SignedTwips;
  right: Twips;
  bottom: SignedTwips;
  left: Twips;
  header: Twips;
  footer: Twips;
  gutter: Twips;
}

/**
 * Source: CT_SectPr (wml.xsd line 1566), EG_SectPrContents
 * Section properties — w:sectPr.
 *
 * Batch 1 additions: pgNumType, textDirection, rtlGutter
 * Batch 2 additions: headerReference, footerReference
 */
export interface WmlSectionProperties {
  type?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  pgSz?: WmlPageSize;
  pgMar?: WmlPageMargins;
  /** Page number format and start value. Source: CT_PageNumber (wml.xsd line 1457) */
  pgNumType?: WmlPageNumbering;
  titlePg?: boolean;            // CT_OnOff — different first page header/footer
  bidi?: boolean;               // CT_OnOff — right-to-left section
  vAlign?: 'top' | 'center' | 'both' | 'bottom';  // CT_VerticalJc
  /** Explicit text direction for the section. Source: CT_TextDirection (wml.xsd line 773) */
  textDirection?: WmlTextDirectionValue;
  /** RTL gutter (gutter on right side for RTL documents). */
  rtlGutter?: boolean;
  /**
   * Header references — link this section to header parts.
   * Source: EG_HdrFtrReferences (wml.xsd line 1522), CT_HdrFtrRef (line 1514)
   * Batch 2 addition.
   */
  headerReference?: import('./header-footer').WmlHeaderFooterRef[];
  /**
   * Footer references — link this section to footer parts.
   * Source: EG_HdrFtrReferences (wml.xsd line 1522), CT_HdrFtrRef (line 1514)
   * Batch 2 addition.
   */
  footerReference?: import('./header-footer').WmlHeaderFooterRef[];
}

// ---------------------------------------------------------------------------
// Body
// Source: CT_Body (wml.xsd line 3233)
// ---------------------------------------------------------------------------

/**
 * Source: CT_Body (wml.xsd line 3233)
 * Document body — w:body.
 * SPEC CONSTRAINT: sectPr must be the last child if present.
 */
export interface WmlBody {
  content: WmlBlockElement[];
  sectPr?: WmlSectionProperties;
}

// ---------------------------------------------------------------------------
// Document
// Source: CT_Document (wml.xsd line 3439)
// ---------------------------------------------------------------------------

/**
 * Source: CT_Document (wml.xsd line 3439)
 * Root element of word/document.xml — w:document.
 */
export interface WmlDocument {
  conformance?: ConformanceClass;
  body: WmlBody;
}

// ---------------------------------------------------------------------------
// Styles
// Source: CT_Styles (wml.xsd ~line 3110), CT_Style (wml.xsd line 3058)
// ---------------------------------------------------------------------------

/**
 * Source: ST_StyleType (wml.xsd ~line 3050)
 */
export type WmlStyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/**
 * Source: CT_Style (wml.xsd line 3058)
 * Style definition — w:style.
 */
export interface WmlStyle {
  type: WmlStyleType;
  styleId: string;
  default?: boolean;
  name: string;
  basedOn?: string;
  next?: string;
  link?: string;
  pPr?: WmlParagraphProperties;
  rPr?: WmlRunProperties;
  /** tblPr / trPr / tcPr for table styles — DEFERRED_V1_1 */
}

/**
 * Source: CT_DocDefaults (wml.xsd line 2687)
 * Document-level default properties — w:docDefaults.
 * Used for consistent formatting across the entire document.
 */
export interface WmlDocDefaults {
  /** Default run properties applied to all runs without explicit rPr. */
  rPrDefault?: WmlRunProperties;
  /** Default paragraph properties applied to all paragraphs without explicit pPr. */
  pPrDefault?: WmlParagraphProperties;
}

/**
 * Source: CT_Styles (wml.xsd ~line 3110)
 * Root of word/styles.xml — w:styles.
 * Batch 5: docDefaults now supported; latentStyles still deferred.
 */
export interface WmlStyles {
  styles: WmlStyle[];
  /** Document-level defaults for runs and paragraphs. */
  docDefaults?: WmlDocDefaults;
}
