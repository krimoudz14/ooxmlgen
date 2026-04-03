/**
 * Document settings — word/settings.xml
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_Settings (line 2733) — root of word/settings.xml
 *
 * Traceability: SPEC_TRACEABILITY.md — DocxSettings, serializeSettings
 *
 * Batch 1 subset: evenAndOddHeaders, updateFields, bidi, defaultTabStop,
 *   themeFontLang (for Arabic/RTL documents).
 * Batch 3 additions: footnotePr, endnotePr (deferred — footnote support).
 *
 * Design: The package exposes generic settings primitives.
 * Applications compose them for their document type (Arabic thesis,
 * English report, etc.). No document-type-specific defaults are hardcoded.
 */

import { el, xmlDocument, XmlElement } from '../shared/xml';
import { NS } from '../shared/namespaces';
import type { Twips, LanguageTag } from '../shared/types';

// ---------------------------------------------------------------------------
// Type definitions
// ---------------------------------------------------------------------------

/**
 * Source: CT_Settings (wml.xsd line 2733) — subset for Batch 1.
 * Document-level settings that affect the entire document.
 */
export interface DocxSettings {
  /**
   * Enable even/odd page header/footer variants.
   * Source: CT_Settings/evenAndOddHeaders (wml.xsd line 2784)
   * Required when using even-page headers (Batch 2).
   */
  evenAndOddHeaders?: boolean;

  /**
   * Force Word to update all fields (PAGE, TOC, SEQ, etc.) when opening.
   * Source: CT_Settings/updateFields (wml.xsd line 2815)
   * Required for TOC and page number fields to display correctly.
   */
  updateFields?: boolean;

  /**
   * Document-level right-to-left layout.
   * Source: CT_Settings — note: bidi is on the document element, not settings,
   * but this flag controls the w:document/@w:bidi attribute behavior.
   * For Arabic/Hebrew/RTL documents.
   */
  bidi?: boolean;

  /**
   * Default tab stop interval in twips.
   * Source: CT_Settings/defaultTabStop (wml.xsd line 2790)
   * Default: 720 twips (0.5 inch).
   */
  defaultTabStop?: Twips;

  /**
   * Theme font language — sets the language for theme font slots.
   * Source: CT_Settings/themeFontLang (wml.xsd line 2822)
   * For Arabic documents: { bidi: 'ar-SA' }
   * For English documents: { val: 'en-US' }
   */
  themeFontLang?: {
    val?: LanguageTag;
    eastAsia?: LanguageTag;
    bidi?: LanguageTag;
  };

  // footnotePr / endnotePr: DEFERRED_BATCH3
}

// ---------------------------------------------------------------------------
// Serializer
// ---------------------------------------------------------------------------

/**
 * Serialize DocxSettings to a complete word/settings.xml string.
 * Source: CT_Settings (wml.xsd line 2733)
 */
export function serializeSettings(settings: DocxSettings): string {
  const children: XmlElement[] = [];

  if (settings.updateFields) {
    children.push(el('w:updateFields', { 'w:val': '1' }));
  }
  if (settings.evenAndOddHeaders) {
    children.push(el('w:evenAndOddHeaders', {}));
  }
  if (settings.defaultTabStop !== undefined) {
    children.push(el('w:defaultTabStop', { 'w:val': settings.defaultTabStop }));
  }
  if (settings.themeFontLang) {
    const attrs: Record<string, string> = {};
    if (settings.themeFontLang.val)      attrs['w:val']      = settings.themeFontLang.val;
    if (settings.themeFontLang.eastAsia) attrs['w:eastAsia'] = settings.themeFontLang.eastAsia;
    if (settings.themeFontLang.bidi)     attrs['w:bidi']     = settings.themeFontLang.bidi;
    children.push(el('w:themeFontLang', attrs));
  }

  const root = el('w:settings', {
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:w':  NS.W,
    'xmlns:r':  NS.R,
  }, ...children);

  return xmlDocument(root);
}
