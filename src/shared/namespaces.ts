/**
 * OOXML XML namespace URIs.
 * Source: ECMA-376 5th ed. - namespace declarations in all schema files.
 *
 * NOTE: Word uses the openxmlformats.org URIs (ISO/IEC 29500 transitional-compatible),
 * not the purl.oclc.org URIs (ECMA strict schema targetNamespace).
 * Both are normatively equivalent per the spec, but Word only accepts the former.
 */

export const NS = {
  // WordprocessingML main namespace
  W: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',

  // DrawingML wordprocessing drawing
  WP: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',

  // DrawingML main
  A: 'http://schemas.openxmlformats.org/drawingml/2006/main',

  // DrawingML picture
  PIC: 'http://schemas.openxmlformats.org/drawingml/2006/picture',

  // Shared relationship references
  R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',

  // Shared common simple types
  S: 'http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes',

  // Office Math
  M: 'http://schemas.openxmlformats.org/officeDocument/2006/math',

  // OPC content types
  CONTENT_TYPES: 'http://schemas.openxmlformats.org/package/2006/content-types',

  // OPC relationships
  RELATIONSHIPS: 'http://schemas.openxmlformats.org/package/2006/relationships',

  // OPC core properties
  CORE_PROPS: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',

  // Extended (app) properties
  EXT_PROPS: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',

  // Dublin Core
  DC: 'http://purl.org/dc/elements/1.1/',
  DCTERMS: 'http://purl.org/dc/terms/',

  // XML namespace
  XML: 'http://www.w3.org/XML/1998/namespace',
  XSI: 'http://www.w3.org/2001/XMLSchema-instance',
} as const;

/**
 * Well-known relationship type URIs.
 * Source: ECMA-376 5th ed. Part 2 (OPC) and Part 1 (WML) relationship type definitions.
 */
export const REL_TYPES = {
  // Package-level
  OFFICE_DOCUMENT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  EXTENDED_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  CUSTOM_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',

  // Document-level
  STYLES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  SETTINGS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  NUMBERING: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  HEADER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  FOOTER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  FOOTNOTES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  ENDNOTES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  COMMENTS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  THEME: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  FONT_TABLE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  WEB_SETTINGS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  IMAGE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  HYPERLINK: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
} as const;

/**
 * Well-known content type strings for DOCX parts.
 * Source: ECMA-376 5th ed. Part 1 - content type definitions.
 */
export const CONTENT_TYPES = {
  DOCX: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  STYLES: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  SETTINGS: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  NUMBERING: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  HEADER: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  FOOTER: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  FOOTNOTES: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  ENDNOTES: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  COMMENTS: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  THEME: 'application/vnd.openxmlformats-officedocument.theme+xml',
  FONT_TABLE: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
  CORE_PROPS: 'application/vnd.openxmlformats-package.core-properties+xml',
  EXT_PROPS: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  RELS: 'application/vnd.openxmlformats-package.relationships+xml',
  XML: 'application/xml',
  PNG: 'image/png',
  JPEG: 'image/jpeg',
  GIF: 'image/gif',
} as const;
