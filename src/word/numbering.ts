/**
 * Numbering definitions — word/numbering.xml
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_Lvl          (line 2961) — one level of a numbering definition
 *   CT_AbstractNum  (line 2989) — abstract (reusable) numbering template
 *   CT_Num          (line 3004) — concrete numbering instance, referenced by paragraphs
 *   CT_Numbering    (line 3015) — root of word/numbering.xml
 *   ST_NumberFormat (line 1294) — number format enum
 *
 * Traceability: SPEC_TRACEABILITY.md — WmlNumberFormat, WmlNumberingLevel,
 *   WmlAbstractNum, WmlNum, WmlNumbering, serializeNumbering
 *
 * Design: The package exposes reusable numbering primitives.
 * Applications compose them into bullet lists, numbered lists, outline
 * numbering, heading numbering, caption numbering, etc.
 * No list style is hardcoded here.
 */

import { el, xmlDocument, XmlElement } from '../shared/xml';
import { NS } from '../shared/namespaces';
import type { WmlParagraphProperties, WmlRunProperties } from './types';
import { serializePPrGeneral, serializeRPr } from './serializer';

// ---------------------------------------------------------------------------
// Type definitions
// ---------------------------------------------------------------------------

/**
 * Source: ST_NumberFormat (wml.xsd line 1294)
 * Number format for list levels and page numbering.
 * Key values for Arabic/multilingual documents included.
 */
export type WmlNumberFormat =
  | 'decimal'
  | 'upperRoman'
  | 'lowerRoman'
  | 'upperLetter'
  | 'lowerLetter'
  | 'ordinal'
  | 'bullet'
  | 'none'
  | 'decimalZero'
  | 'arabicAbjad'
  | 'arabicAlpha'
  | 'hebrew1'
  | 'hebrew2'
  | 'numberInDash'
  | 'cardinalText'
  | 'ordinalText'
  | string; // allow any valid ST_NumberFormat value

/**
 * Source: ST_LevelSuffix (wml.xsd ~line 2950)
 * Character after the level number.
 */
export type WmlLevelSuffix = 'tab' | 'space' | 'nothing';

/**
 * Source: CT_Lvl (wml.xsd line 2961)
 * One level of a numbering definition (0-based, 0-8).
 */
export interface WmlNumberingLevel {
  /** Level index 0-8. Required. */
  ilvl: number;
  /** Starting number. Default 1. */
  start?: number;
  /** Number format. Default 'decimal'. */
  numFmt?: WmlNumberFormat;
  /** Level text pattern. Use %1, %2 etc. for level numbers. e.g. "%1." or "•" */
  lvlText?: string;
  /** Alignment of the number itself. */
  lvlJc?: 'start' | 'center' | 'end';
  /** Character after the number: tab (default), space, or nothing. */
  suff?: WmlLevelSuffix;
  /** Restart level numbering at this value (0 = never restart). */
  lvlRestart?: number;
  /** Link this level to a paragraph style ID (e.g. 'Heading1'). */
  pStyle?: string;
  /** Paragraph formatting for this level (indent, spacing, etc.). */
  pPr?: WmlParagraphProperties;
  /** Run formatting for the number itself (font, size, bold, etc.). */
  rPr?: WmlRunProperties;
}

/**
 * Source: CT_AbstractNum (wml.xsd line 2989)
 * Abstract (reusable) numbering template. Referenced by CT_Num.
 */
export interface WmlAbstractNum {
  /** Auto-assigned if omitted. */
  abstractNumId?: number;
  /** 'singleLevel' | 'multilevel' | 'hybridMultilevel'. Default 'multilevel'. */
  multiLevelType?: 'singleLevel' | 'multilevel' | 'hybridMultilevel';
  /** Optional name for this abstract numbering (informational). */
  name?: string;
  /** Level definitions. Up to 9 levels (ilvl 0-8). */
  levels: WmlNumberingLevel[];
}

/**
 * Source: CT_Num (wml.xsd line 3004)
 * Concrete numbering instance. Paragraphs reference this via w:numPr/w:numId.
 */
export interface WmlNum {
  /** Auto-assigned if omitted. */
  numId?: number;
  /** References WmlAbstractNum.abstractNumId. */
  abstractNumId: number;
}

/**
 * Source: CT_Numbering (wml.xsd line 3015)
 * Root of word/numbering.xml.
 */
export interface WmlNumbering {
  abstractNums: WmlAbstractNum[];
  nums: WmlNum[];
}

// ---------------------------------------------------------------------------
// Serializer
// ---------------------------------------------------------------------------

function serializeLvl(lvl: WmlNumberingLevel): XmlElement {
  const children: XmlElement[] = [];

  if (lvl.start !== undefined) {
    children.push(el('w:start', { 'w:val': lvl.start }));
  }
  if (lvl.numFmt) {
    children.push(el('w:numFmt', { 'w:val': lvl.numFmt }));
  }
  if (lvl.lvlRestart !== undefined) {
    children.push(el('w:lvlRestart', { 'w:val': lvl.lvlRestart }));
  }
  if (lvl.pStyle) {
    children.push(el('w:pStyle', { 'w:val': lvl.pStyle }));
  }
  if (lvl.suff) {
    children.push(el('w:suff', { 'w:val': lvl.suff }));
  }
  if (lvl.lvlText !== undefined) {
    children.push(el('w:lvlText', { 'w:val': lvl.lvlText }));
  }
  if (lvl.lvlJc) {
    children.push(el('w:lvlJc', { 'w:val': lvl.lvlJc }));
  }
  if (lvl.pPr) {
    children.push(serializePPrGeneral(lvl.pPr));
  }
  if (lvl.rPr) {
    children.push(serializeRPr(lvl.rPr));
  }

  return el('w:lvl', { 'w:ilvl': lvl.ilvl }, ...children);
}

function serializeAbstractNum(an: WmlAbstractNum): XmlElement {
  const children: XmlElement[] = [];

  if (an.multiLevelType) {
    children.push(el('w:multiLevelType', { 'w:val': an.multiLevelType }));
  }
  if (an.name) {
    children.push(el('w:name', { 'w:val': an.name }));
  }
  for (const lvl of an.levels) {
    children.push(serializeLvl(lvl));
  }

  return el('w:abstractNum', { 'w:abstractNumId': an.abstractNumId! }, ...children);
}

function serializeNum(num: WmlNum): XmlElement {
  return el('w:num', { 'w:numId': num.numId! },
    el('w:abstractNumId', { 'w:val': num.abstractNumId }),
  );
}

/**
 * Serialize a WmlNumbering to a complete word/numbering.xml string.
 * Source: CT_Numbering (wml.xsd line 3015)
 */
export function serializeNumbering(numbering: WmlNumbering): string {
  const children: XmlElement[] = [
    ...numbering.abstractNums.map(serializeAbstractNum),
    ...numbering.nums.map(serializeNum),
  ];

  const root = el('w:numbering', {
    'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:w':   NS.W,
    'xmlns:r':   NS.R,
  }, ...children);

  return xmlDocument(root);
}

// ---------------------------------------------------------------------------
// Numbering manager — used by DocxBuilder to assign IDs and track state
// ---------------------------------------------------------------------------

export class NumberingManager {
  private abstractNums: WmlAbstractNum[] = [];
  private nums: WmlNum[] = [];
  private nextAbstractId = 0;
  private nextNumId = 1;

  /**
   * Define a new numbering. Returns the numId to use in w:numPr.
   * Source: CT_AbstractNum + CT_Num pattern
   */
  define(def: Omit<WmlAbstractNum, 'abstractNumId'>): number {
    const abstractNumId = this.nextAbstractId++;
    this.abstractNums.push({ ...def, abstractNumId });

    const numId = this.nextNumId++;
    this.nums.push({ numId, abstractNumId });

    return numId;
  }

  hasDefinitions(): boolean {
    return this.nums.length > 0;
  }

  toWmlNumbering(): WmlNumbering {
    return {
      abstractNums: this.abstractNums,
      nums: this.nums,
    };
  }
}

// ---------------------------------------------------------------------------
// Preset numbering helpers — reusable starting points, not hardcoded styles
// ---------------------------------------------------------------------------

/**
 * Bullet list — single level, bullet character.
 * Application can override any property.
 */
export function bulletListLevel(ilvl = 0, overrides: Partial<WmlNumberingLevel> = {}): WmlNumberingLevel {
  return {
    ilvl,
    start: 1,
    numFmt: 'bullet',
    lvlText: '•',
    lvlJc: 'start',
    suff: 'tab',
    pPr: { ind: { start: 720 * (ilvl + 1), hanging: 360 } },
    rPr: { rFonts: { ascii: 'Symbol', hAnsi: 'Symbol', cs: 'Symbol' } },
    ...overrides,
  };
}

/**
 * Decimal numbered list — single level.
 * Application can override any property.
 */
export function decimalListLevel(ilvl = 0, overrides: Partial<WmlNumberingLevel> = {}): WmlNumberingLevel {
  return {
    ilvl,
    start: 1,
    numFmt: 'decimal',
    lvlText: `%${ilvl + 1}.`,
    lvlJc: 'start',
    suff: 'tab',
    pPr: { ind: { start: 720 * (ilvl + 1), hanging: 360 } },
    ...overrides,
  };
}

/**
 * Arabic Abjad list level — for Arabic documents.
 * Application can override any property.
 */
export function arabicAbjadListLevel(ilvl = 0, overrides: Partial<WmlNumberingLevel> = {}): WmlNumberingLevel {
  return {
    ilvl,
    start: 1,
    numFmt: 'arabicAbjad',
    lvlText: `%${ilvl + 1}.`,
    lvlJc: 'start',
    suff: 'tab',
    pPr: { ind: { start: 720 * (ilvl + 1), hanging: 360 }, bidi: true },
    ...overrides,
  };
}
