/**
 * Footnotes and Endnotes — word/footnotes.xml, word/endnotes.xml
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_FtnEdn      (line 2491) — individual footnote/endnote entry
 *   CT_FtnEdnRef   (line 2484) — inline reference mark in document body
 *   CT_Footnotes   (line 3245) — root of word/footnotes.xml
 *   CT_Endnotes    (line 3253) — root of word/endnotes.xml
 *   CT_FtnProps    (line 2506) — footnote properties (position, numFmt, restart)
 *   CT_EdnProps    (line 2513) — endnote properties
 *   ST_FtnEdn      (line ~2470) — normal | separator | continuationSeparator | continuationNotice
 *
 * Traceability: SPEC_TRACEABILITY.md — WmlFootnote, WmlFootnoteRef,
 *   serializeFootnotes, serializeEndnotes
 *
 * Key spec constraints:
 * - footnotes.xml MUST contain separator (id=-1) and continuationSeparator (id=0)
 * - endnotes.xml MUST contain separator (id=-1) and continuationSeparator (id=0)
 * - User notes start at id=1, auto-incremented
 * - w:footnoteReference/@w:id must match CT_FtnEdn/@w:id in footnotes.xml
 * - Content type: application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml
 *
 * Design: The package exposes generic note primitives.
 * Applications decide note content (citations, explanations, legal references).
 * No citation style or bibliography format is hardcoded.
 */

import { el, xmlDocument, XmlElement } from '../shared/xml';
import { NS } from '../shared/namespaces';
import type { WmlBlockElement, WmlRun } from './types';
import type { WmlNumberFormat } from './numbering';
import { serializeBlock } from './serializer';

// ---------------------------------------------------------------------------
// Type definitions
// ---------------------------------------------------------------------------

/**
 * Source: ST_FtnEdn (wml.xsd ~line 2470)
 * Type of a footnote/endnote entry.
 * 'normal' = user note; 'separator' and 'continuationSeparator' are required
 * system entries that Word uses to draw the separator line.
 */
export type WmlNoteType = 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice';

/**
 * Source: CT_FtnEdn (wml.xsd line 2491)
 * Individual footnote or endnote entry.
 */
export interface WmlFootnote {
  /** Note ID. -1 = separator, 0 = continuationSeparator, 1+ = user notes. */
  id: number;
  /** Note type. Default 'normal' for user notes. */
  type?: WmlNoteType;
  /** Block-level content of the note. Same model as document body. */
  content: WmlBlockElement[];
}

/**
 * Source: CT_FtnEdnRef (wml.xsd line 2484)
 * Inline reference mark placed in a run in the document body.
 * Used as run content (_type: 'footnoteReference' | 'endnoteReference').
 */
export interface WmlFootnoteRef {
  _type: 'footnoteReference' | 'endnoteReference';
  /** Must match the id of the corresponding WmlFootnote in footnotes/endnotes.xml. */
  id: number;
  /** If true, a custom mark follows this reference (rare). */
  customMarkFollows?: boolean;
}

/**
 * Source: CT_FtnProps (wml.xsd line 2506)
 * Footnote properties — controls position, number format, and restart behavior.
 * Used in sectPr.footnotePr and settings.footnotePr.
 */
export interface WmlFootnoteProperties {
  /** Position: 'pageBottom' (default) | 'beneathText' | 'sectEnd' | 'docEnd' */
  pos?: 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';
  /** Number format for footnote marks. Default 'decimal'. */
  numFmt?: WmlNumberFormat;
  /** Starting number. Default 1. */
  numStart?: number;
  /** When to restart numbering. Default 'continuous'. */
  numRestart?: 'continuous' | 'eachSect' | 'eachPage';
}

/**
 * Source: CT_EdnProps (wml.xsd line 2513)
 * Endnote properties — same structure as footnote properties.
 */
export interface WmlEndnoteProperties {
  /** Position: 'sectEnd' (default) | 'docEnd' */
  pos?: 'sectEnd' | 'docEnd';
  numFmt?: WmlNumberFormat;
  numStart?: number;
  numRestart?: 'continuous' | 'eachSect' | 'eachPage';
}

// ---------------------------------------------------------------------------
// Required system entries
// Source: CT_FtnDocProps (wml.xsd line 2519) — separator and continuationSeparator
// are required in every footnotes.xml and endnotes.xml per spec.
// ---------------------------------------------------------------------------

/**
 * The required separator footnote (id=-1).
 * Source: CT_FtnEdn type="separator"
 * Word uses this to draw the horizontal line above footnotes.
 */
const SEPARATOR_NOTE: WmlFootnote = {
  id: -1,
  type: 'separator',
  content: [{ _type: 'paragraph', content: [] }],
};

/**
 * The required continuation separator footnote (id=0).
 * Source: CT_FtnEdn type="continuationSeparator"
 * Word uses this when footnotes continue from the previous page.
 */
const CONTINUATION_SEPARATOR_NOTE: WmlFootnote = {
  id: 0,
  type: 'continuationSeparator',
  content: [{ _type: 'paragraph', content: [] }],
};

// ---------------------------------------------------------------------------
// Serializers
// ---------------------------------------------------------------------------

function serializeNote(note: WmlFootnote, elementName: 'footnote' | 'endnote'): XmlElement {
  const attrs: Record<string, string | number> = { 'w:id': note.id };
  if (note.type && note.type !== 'normal') {
    attrs['w:type'] = note.type;
  }
  const children = note.content.map(serializeBlock);
  return el(`w:${elementName}`, attrs, ...children);
}

function serializeFtnProps(props: WmlFootnoteProperties, elementName: string): XmlElement {
  const children: XmlElement[] = [];
  if (props.pos)        children.push(el('w:pos',        { 'w:val': props.pos }));
  if (props.numFmt)     children.push(el('w:numFmt',     { 'w:val': props.numFmt }));
  if (props.numStart !== undefined) children.push(el('w:numStart', { 'w:val': props.numStart }));
  if (props.numRestart) children.push(el('w:numRestart', { 'w:val': props.numRestart }));
  return el(elementName, {}, ...children);
}

/**
 * Serialize footnotes to a complete word/footnotes.xml string.
 * Source: CT_Footnotes (wml.xsd line 3245)
 * Always includes the required separator and continuationSeparator entries.
 */
export function serializeFootnotes(
  userNotes: WmlFootnote[],
  props?: WmlFootnoteProperties,
): string {
  const allNotes = [SEPARATOR_NOTE, CONTINUATION_SEPARATOR_NOTE, ...userNotes];
  const children: XmlElement[] = allNotes.map(n => serializeNote(n, 'footnote'));

  const attrs: Record<string, string> = {
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:r':   NS.R,
    'xmlns:m':   NS.M,
    'xmlns:w':   NS.W,
  };

  const root = el('w:footnotes', attrs, ...children);
  return xmlDocument(root);
}

/**
 * Serialize endnotes to a complete word/endnotes.xml string.
 * Source: CT_Endnotes (wml.xsd line 3253)
 * Always includes the required separator and continuationSeparator entries.
 */
export function serializeEndnotes(
  userNotes: WmlFootnote[],
  props?: WmlEndnoteProperties,
): string {
  const allNotes = [SEPARATOR_NOTE, CONTINUATION_SEPARATOR_NOTE, ...userNotes];
  const children: XmlElement[] = allNotes.map(n => serializeNote(n, 'endnote'));

  const attrs: Record<string, string> = {
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:r':   NS.R,
    'xmlns:m':   NS.M,
    'xmlns:w':   NS.W,
  };

  const root = el('w:endnotes', attrs, ...children);
  return xmlDocument(root);
}

// ---------------------------------------------------------------------------
// Note manager — used by DocxBuilder to track notes and assign IDs
// ---------------------------------------------------------------------------

export class NoteManager {
  private footnotes: WmlFootnote[] = [];
  private endnotes: WmlFootnote[] = [];
  private nextFootnoteId = 1;
  private nextEndnoteId = 1;
  private footnoteProps?: WmlFootnoteProperties;
  private endnoteProps?: WmlEndnoteProperties;

  /**
   * Add a footnote and return a WmlFootnoteRef run content item.
   * The returned ref must be placed inside a WmlRun in the document body.
   */
  addFootnote(content: WmlBlockElement[]): WmlFootnoteRef {
    const id = this.nextFootnoteId++;
    this.footnotes.push({ id, type: 'normal', content });
    return { _type: 'footnoteReference', id };
  }

  /**
   * Add an endnote and return a WmlFootnoteRef run content item.
   */
  addEndnote(content: WmlBlockElement[]): WmlFootnoteRef {
    const id = this.nextEndnoteId++;
    this.endnotes.push({ id, type: 'normal', content });
    return { _type: 'endnoteReference', id };
  }

  setFootnoteProperties(props: WmlFootnoteProperties): void {
    this.footnoteProps = props;
  }

  setEndnoteProperties(props: WmlEndnoteProperties): void {
    this.endnoteProps = props;
  }

  hasFootnotes(): boolean { return this.footnotes.length > 0; }
  hasEndnotes(): boolean  { return this.endnotes.length > 0; }

  getFootnotes(): WmlFootnote[]          { return this.footnotes; }
  getEndnotes(): WmlFootnote[]           { return this.endnotes; }
  getFootnoteProps(): WmlFootnoteProperties | undefined { return this.footnoteProps; }
  getEndnoteProps(): WmlEndnoteProperties | undefined   { return this.endnoteProps; }
}
