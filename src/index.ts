/**
 * ooxmlgen — OOXML generation library
 * v1: DOCX / WordprocessingML (Strict, ECMA-376 5th edition)
 *
 * PRIMARY API: DocxBuilder + factory helpers + measurement utilities.
 * ADVANCED API: Low-level serializers, OPC classes, XML builder.
 *
 * V1 SCOPE:
 *   - OPC package (content types, relationships, ZIP)
 *   - Paragraphs, runs, text, breaks, tabs
 *   - Run formatting: bold, italic, underline, color, font, size, etc.
 *   - Paragraph formatting: alignment, spacing, indentation, outline level
 *   - Tables: borders, column widths, cell shading, column spanning
 *   - Styles: paragraph and character style definitions
 *   - Section properties: page size, margins, orientation
 *   - Core and extended document properties
 *
 * DEFERRED (not exported):
 *   - Numbering / lists          → v1.1
 *   - Images / drawing           → v1.1
 *   - Hyperlinks                 → v1.1
 *   - Headers / footers          → v1.1
 *   - Footnotes / endnotes       → v1.1
 *   - Fields (page numbers, TOC) → v1.1
 *   - Document settings          → v1.1
 *   - Font table                 → v1.1
 *   - XLSX / SpreadsheetML       → v2
 *   - PPTX / PresentationML      → v3
 */

// ---------------------------------------------------------------------------
// PRIMARY API — DocxBuilder and factory helpers
// ---------------------------------------------------------------------------

export { DocxBuilder } from './word/builder';
export { paragraph, textRun, emptyParagraph, pageBreak } from './word/builder';

// Measurement helpers
export { ptToTwips, inToTwips, cmToTwips, ptToHalfPt } from './word/builder';

// Document property types
export type { DocxCoreProperties, DocxExtendedProperties } from './word/builder';

// ---------------------------------------------------------------------------
// BATCH 1 — Numbering, Settings, Page Numbering
// ---------------------------------------------------------------------------

// Numbering definition API
export { serializeNumbering, bulletListLevel, decimalListLevel, arabicAbjadListLevel } from './word/numbering';
export type {
  WmlNumberFormat, WmlNumberingLevel, WmlAbstractNum, WmlNum, WmlNumbering,
  WmlLevelSuffix,
} from './word/numbering';

// Settings API
export { serializeSettings } from './word/settings';
export type { DocxSettings } from './word/settings';

// ---------------------------------------------------------------------------
// BATCH 2 — Headers, Footers, Page Number Fields
// ---------------------------------------------------------------------------

// Header/footer API
export { serializeHeader, serializeFooter } from './word/header-footer';
export type { WmlHeaderFooterType, WmlHeaderFooterRef } from './word/header-footer';

// Field infrastructure
export { fieldRuns, pageNumberField, totalPagesField, pageXofYRuns } from './word/fields';
export type { WmlFldChar, WmlInstrText } from './word/fields';

// ---------------------------------------------------------------------------
// BATCH 3 — Footnotes and Endnotes
// ---------------------------------------------------------------------------

export { serializeFootnotes, serializeEndnotes } from './word/footnotes';
export type {
  WmlNoteType,
  WmlFootnote,
  WmlFootnoteRef,
  WmlFootnoteProperties,
  WmlEndnoteProperties,
} from './word/footnotes';

// ---------------------------------------------------------------------------
// BATCH 4 — Hyperlinks, TOC, SEQ/caption fields
// ---------------------------------------------------------------------------

export { createHyperlink } from './word/hyperlinks';
export {
  tocField,
  seqField,
  styleRefField,
  refField,
} from './word/fields';
export {
  createCaptionParagraph,
  createTableOfContents,
} from './word/helpers';
export type { WmlHyperlink } from './word/hyperlinks';

// ---------------------------------------------------------------------------
// CONTENT TYPES — used when constructing paragraphs, runs, tables, styles
// ---------------------------------------------------------------------------

export type {
  // Block-level
  WmlBlockElement,
  WmlParagraph,
  WmlParagraphProperties,
  WmlTable,
  WmlTableRow,
  WmlTableCell,
  WmlTableProperties,
  WmlTableGrid,
  WmlTableRowProperties,
  WmlTableCellProperties,  // Batch 1: now fully supported
  WmlTableBorders,
  WmlTableWidth,
  WmlTableCellBorders,     // Batch 1 addition
  WmlTableCellMargins,     // Batch 1 addition

  // Inline
  WmlRun,
  WmlRunProperties,
  WmlText,
  WmlBreak,

  // Section
  WmlSectionProperties,
  WmlPageSize,
  WmlPageMargins,
  WmlPageNumbering,        // Batch 1 addition
  WmlTextDirectionValue,   // Batch 1 addition

  // Styles
  WmlStyle,
  WmlStyleType,
  WmlDocDefaults,

  // Formatting sub-types
  WmlColor,
  WmlFonts,
  WmlSpacing,
  WmlIndentation,
  WmlBorder,
  WmlShading,
} from './word/types';

// Measurement primitives (defined in shared/types)
export type { Twips, HalfPoints, HexColor, SignedTwips } from './shared/types';

// NOTE: The following types are intentionally NOT exported from the public API:
//   WmlDocument, WmlBody, WmlStyles — internal to DocxBuilder; consumers never
//     construct these directly.
//   WmlNumberingProperties — PROVISIONAL/DEFERRED: field exists on
//     WmlParagraphProperties.numPr but the serializer silently drops it until
//     word/numbering.xml support is added in v1.1.
//   WmlTab, WmlCarriageReturn — internal run content; use the WmlRunContent
//     union via WmlRun.content if needed.

// ---------------------------------------------------------------------------
// ADVANCED API — low-level access for custom serialization or OPC manipulation
// ---------------------------------------------------------------------------

/** @advanced Serialize a WmlDocument object to a word/document.xml string. */
export { serializeDocument, serializeStyles } from './word/serializer';

/** @advanced OPC package builder — wraps JSZip. */
export { OpcPackage } from './core/package';
export { ContentTypeManager } from './core/content-types';
export { RelationshipManager } from './core/relationships';
export type { OpcRelationship, OpcTargetMode } from './core/relationships';

/** @advanced Namespace URI constants. */
export { NS, REL_TYPES, CONTENT_TYPES } from './shared/namespaces';

/** @advanced Minimal XML element builder. */
export { el, serialize, xmlDocument } from './shared/xml';
export type { XmlElement, XmlAttrs, XmlNode } from './shared/xml';
