/**
 * Field infrastructure — fldChar / instrText run content.
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_FldChar     (line 1210) — field character marker
 *   ST_FldCharType (line 1175) — begin / separate / end
 *   instrText      (EG_RunInnerContent) — field instruction string
 *
 * Traceability: SPEC_TRACEABILITY.md — WmlFldChar, WmlInstrText,
 *   pageNumberField, totalPagesField, fieldRuns
 *
 * Batch 2 scope: PAGE and NUMPAGES fields for use in headers/footers.
 * Batch 4 will extend this with TOC, SEQ, hyperlinks.
 *
 * Design: The package exposes generic field primitives.
 * Applications compose them into page numbers, TOC entries, captions, etc.
 * No field instruction is hardcoded as a document-type-specific concept.
 *
 * Field character sequence (from CT_FldChar spec):
 *   w:r > w:fldChar fldCharType="begin"
 *   w:r > w:instrText (xml:space="preserve") " PAGE "
 *   w:r > w:fldChar fldCharType="separate"
 *   w:r > w:t  (cached display value — "1" by default)
 *   w:r > w:fldChar fldCharType="end"
 */

import type { WmlRun, WmlRunProperties } from './types';

// ---------------------------------------------------------------------------
// Run content types — extend WmlRunContent union in Batch 2
// ---------------------------------------------------------------------------

/**
 * Source: CT_FldChar (wml.xsd line 1210)
 * Field character — marks begin, separate, or end of a complex field.
 */
export interface WmlFldChar {
  _type: 'fldChar';
  fldCharType: 'begin' | 'separate' | 'end';
  fldLock?: boolean;
  dirty?: boolean;
}

/**
 * Source: instrText element in EG_RunInnerContent (wml.xsd line 1692)
 * Field instruction text — carries the field instruction string.
 * xml:space="preserve" is always set (instruction strings have spaces).
 */
export interface WmlInstrText {
  _type: 'instrText';
  text: string;
}

// ---------------------------------------------------------------------------
// Factory helpers
// ---------------------------------------------------------------------------

/**
 * Build a complete complex field run sequence.
 * Source: CT_FldChar begin/separate/end pattern (wml.xsd line 1210)
 *
 * @param instruction - Field instruction string, e.g. " PAGE ", " NUMPAGES "
 * @param displayText - Cached display value shown before Word updates the field
 * @param rPr - Optional run properties applied to all runs in the sequence
 */
export function fieldRuns(
  instruction: string,
  displayText = '1',
  rPr?: WmlRunProperties,
): WmlRun[] {
  return [
    // begin
    { _type: 'run', rPr, content: [{ _type: 'fldChar', fldCharType: 'begin' } as WmlFldChar] },
    // instruction
    { _type: 'run', rPr, content: [{ _type: 'instrText', text: instruction } as WmlInstrText] },
    // separate
    { _type: 'run', rPr, content: [{ _type: 'fldChar', fldCharType: 'separate' } as WmlFldChar] },
    // cached display value
    { _type: 'run', rPr, content: [{ _type: 'text', text: displayText, space: 'preserve' }] },
    // end
    { _type: 'run', rPr, content: [{ _type: 'fldChar', fldCharType: 'end' } as WmlFldChar] },
  ];
}

/**
 * PAGE field — current page number.
 * Source: PAGE field instruction (ECMA-376 Part1 §17.16.5.45)
 */
export function pageNumberField(rPr?: WmlRunProperties): WmlRun[] {
  return fieldRuns(' PAGE ', '1', rPr);
}

/**
 * NUMPAGES field — total page count.
 * Source: NUMPAGES field instruction (ECMA-376 Part1 §17.16.5.43)
 */
export function totalPagesField(rPr?: WmlRunProperties): WmlRun[] {
  return fieldRuns(' NUMPAGES ', '1', rPr);
}

/**
 * "Page X of Y" run sequence — combines PAGE and NUMPAGES.
 * Returns an array of runs: [text "Page "] + PAGE field + [text " of "] + NUMPAGES field
 */
export function pageXofYRuns(rPr?: WmlRunProperties): WmlRun[] {
  const textRun = (text: string): WmlRun => ({
    _type: 'run',
    rPr,
    content: [{ _type: 'text', text, space: 'preserve' }],
  });
  return [
    textRun('Page '),
    ...pageNumberField(rPr),
    textRun(' of '),
    ...totalPagesField(rPr),
  ];
}

// ---------------------------------------------------------------------------
// Batch 4 additions — TOC, SEQ, caption fields
// ---------------------------------------------------------------------------

/**
 * TOC (Table of Contents) field instruction.
 * Source: TOC field instruction (ECMA-376 Part1 §17.16.5.55)
 *
 * @param switches - Optional TOC switches, e.g. '\\o "1-3" \\h \\z \\u'
 * @returns A complete TOC field run sequence
 *
 * @example
 * // Basic TOC with heading levels 1-3
 * const tocRuns = tocField({ headingLevels: [1, 2, 3], withHyperlinks: true });
 * builder.addBlock(paragraph(tocRuns));
 *
 * // Custom TOC instruction
 * const tocRuns = tocField({ customSwitches: '\\o "1-3" \\h \\z \\u' });
 * builder.addBlock(paragraph(tocRuns));
 */
export function tocField(options: {
  headingLevels?: number[];
  withHyperlinks?: boolean;
  withPageNumbers?: boolean;
  preserveTabs?: boolean;
  customSwitches?: string;
  displayText?: string;
  rPr?: WmlRunProperties;
} = {}): WmlRun[] {
  // Build the instruction string
  const parts: string[] = ['TOC'];

  if (options.customSwitches) {
    parts.push(options.customSwitches);
  } else {
    if (options.headingLevels) {
      const range = `${Math.min(...options.headingLevels)}-${Math.max(...options.headingLevels)}`;
      parts.push(`\\o "${range}"`);
    }
    if (options.withHyperlinks) parts.push('\\h');
    if (options.withPageNumbers !== false) parts.push('\\z');
    if (options.preserveTabs) parts.push('\\u');
  }

  const instruction = ` ${parts.join(' ')} `;
  const displayText = options.displayText || 'Table of Contents';
  return fieldRuns(instruction, displayText, options.rPr);
}

/**
 * SEQ (Sequence) field — used for numbering figures, tables, equations, etc.
 * Source: SEQ field instruction (ECMA-376 Part1 §17.16.5.52)
 *
 * @param identifier - The sequence name, e.g. "Figure", "Table", "Equation"
 * @param switches - Optional switches: \\c (nearest preceding), \\h, \\n, \\r N, \\s N
 * @returns A complete SEQ field run sequence
 *
 * @example
 * // "Figure 1"
 * const seqRuns = seqField('Figure');
 * builder.addBlock(paragraph([{ text: 'Figure ' }, ...seqRuns, ': My caption']));
 *
 * // Reset sequence to 5
 * const seqRuns = seqField('Table', { resetTo: 5 });
 */
export function seqField(
  identifier: string,
  options: {
    resetTo?: number;
    restartAtHeading?: number;
    nearestPreceding?: boolean;
    hideInTOC?: boolean;
    displayText?: string;
    rPr?: WmlRunProperties;
  } = {},
): WmlRun[] {
  const parts: string[] = [`SEQ ${identifier}`];

  if (options.resetTo !== undefined) parts.push(`\\r ${options.resetTo}`);
  if (options.restartAtHeading !== undefined) parts.push(`\\s ${options.restartAtHeading}`);
  if (options.nearestPreceding) parts.push('\\c');
  if (options.hideInTOC) parts.push('\\h');

  const instruction = ` ${parts.join(' ')} `;
  const displayText = options.displayText || '1';
  return fieldRuns(instruction, displayText, options.rPr);
}

/**
 * STYLEREF field — reference a style's text (used in headers for running headers).
 * Source: STYLEREF field instruction (ECMA-376 Part1 §17.16.5.54)
 *
 * @param styleId - The style ID to reference, e.g. "Heading1"
 * @returns A complete STYLEREF field run sequence
 */
export function styleRefField(
  styleId: string,
  options: {
    firstMatch?: boolean;
    nearestPreceding?: boolean;
    displayText?: string;
    rPr?: WmlRunProperties;
  } = {},
): WmlRun[] {
  const parts: string[] = [`STYLEREF "${styleId}"`];
  if (options.firstMatch) parts.push('\\t');
  if (options.nearestPreceding) parts.push('\\n');

  const instruction = ` ${parts.join(' ')} `;
  const displayText = options.displayText || '';
  return fieldRuns(instruction, displayText, options.rPr);
}

/**
 * REF field — cross-reference to a bookmark.
 * Source: REF field instruction (ECMA-376 Part1 §17.16.5.47)
 *
 * @param bookmarkName - The bookmark name to reference
 * @returns A complete REF field run sequence
 */
export function refField(
  bookmarkName: string,
  options: {
    insertParagraphNumber?: boolean;
    insertRelativePosition?: boolean;
    displayText?: string;
    rPr?: WmlRunProperties;
  } = {},
): WmlRun[] {
  const parts: string[] = [`REF ${bookmarkName}`];
  if (options.insertParagraphNumber) parts.push('\\p');
  if (options.insertRelativePosition) parts.push('\\r');

  const instruction = ` ${parts.join(' ')} `;
  const displayText = options.displayText || '';
  return fieldRuns(instruction, displayText, options.rPr);
}

// ---------------------------------------------------------------------------
// Batch 4 caption helpers
// ---------------------------------------------------------------------------

// Note: captionParagraph and tableOfContents moved to helpers.ts to avoid circular dependencies
