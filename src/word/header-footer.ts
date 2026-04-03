/**
 * Header and footer parts — word/header*.xml, word/footer*.xml
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_HdrFtr     (line 1528) — content model: EG_BlockLevelElts (same as body)
 *   CT_HdrFtrRef  (line 1514) — links section to header/footer part via r:id + type
 *   ST_HdrFtr     (line 1510) — 'default' | 'first' | 'even'
 *   EG_HdrFtrReferences (line 1522) — headerReference / footerReference in sectPr
 *
 * Traceability: SPEC_TRACEABILITY.md — WmlHeaderFooterType, WmlHeaderFooterRef,
 *   serializeHeader, serializeFooter
 *
 * Design: The package provides generic header/footer part serialization.
 * Applications decide what content goes in headers/footers.
 * No header/footer content is hardcoded.
 *
 * Key spec constraints:
 * - Header/footer parts are separate XML files with their own relationships
 * - sectPr.headerReference / footerReference link sections to parts by rId + type
 * - titlePg=true in sectPr enables the 'first' page variant
 * - evenAndOddHeaders=true in settings enables the 'even' page variant
 * - Content type: application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml
 */

import { el, xmlDocument } from '../shared/xml';
import { NS } from '../shared/namespaces';
import type { WmlBlockElement } from './types';
import { serializeBlock } from './serializer';

// ---------------------------------------------------------------------------
// Type definitions
// ---------------------------------------------------------------------------

/**
 * Source: ST_HdrFtr (wml.xsd line 1510)
 * Header/footer variant type.
 * - 'default': used on all pages unless overridden by 'first' or 'even'
 * - 'first': first page of the section (requires titlePg=true in sectPr)
 * - 'even': even-numbered pages (requires evenAndOddHeaders=true in settings)
 */
export type WmlHeaderFooterType = 'default' | 'first' | 'even';

/**
 * Source: CT_HdrFtrRef (wml.xsd line 1514)
 * Links a section to a header or footer part.
 * Used in sectPr.headerReference[] and sectPr.footerReference[].
 */
export interface WmlHeaderFooterRef {
  type: WmlHeaderFooterType;
  /** Relationship ID — references the header/footer part. */
  id: string;
}

// ---------------------------------------------------------------------------
// Serializers
// ---------------------------------------------------------------------------

/**
 * Serialize header content to a complete word/header*.xml string.
 * Source: CT_HdrFtr (wml.xsd line 1528) — root element w:hdr
 * Content model: EG_BlockLevelElts (same as document body)
 */
export function serializeHeader(content: WmlBlockElement[]): string {
  const children = content.map(serializeBlock);
  const root = el('w:hdr', {
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:r':   NS.R,
    'xmlns:m':   NS.M,
    'xmlns:w':   NS.W,
  }, ...children);
  return xmlDocument(root);
}

/**
 * Serialize footer content to a complete word/footer*.xml string.
 * Source: CT_HdrFtr (wml.xsd line 1528) — root element w:ftr
 * Content model: EG_BlockLevelElts (same as document body)
 */
export function serializeFooter(content: WmlBlockElement[]): string {
  const children = content.map(serializeBlock);
  const root = el('w:ftr', {
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:r':   NS.R,
    'xmlns:m':   NS.M,
    'xmlns:w':   NS.W,
  }, ...children);
  return xmlDocument(root);
}
