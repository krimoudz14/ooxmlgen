/**
 * DocxBuilder — high-level API for constructing DOCX documents.
 *
 * Source: ECMA-376 5th ed.
 *   Part1/wml.xsd (partially read — see repo-manifest.json)
 *   Part2/opc-contentTypes.xsd (fully read)
 *   Part2/opc-relationships.xsd (fully read)
 *   Part2/opc-coreProperties.xsd (fully read)
 *   Part1/shared-documentPropertiesExtended.xsd (fully read)
 * Traceability: SPEC_TRACEABILITY.md — DocxBuilder
 *
 * Batch 1 additions:
 *   - word/numbering.xml    (defineNumbering)
 *   - word/settings.xml     (configureSettings)
 *   - pgNumType on sections
 *   - textDirection on sections and paragraphs
 *   - full table cell properties (tcBorders, vMerge, noWrap, tcMar)
 *
 * DEFERRED parts (not yet produced):
 *   - word/header*.xml      → DEFERRED_BATCH2
 *   - word/footer*.xml      → DEFERRED_BATCH2
 *   - word/footnotes.xml    → DEFERRED_BATCH3
 *   - word/endnotes.xml     → DEFERRED_BATCH3
 *   - word/fontTable.xml    → DEFERRED_BATCH5
 *   - word/comments.xml     → DEFERRED_V2
 */

import { OpcPackage } from '../core/package';
import { REL_TYPES, CONTENT_TYPES, NS } from '../shared/namespaces';
import { el, xmlDocument } from '../shared/xml';
import { serializeDocument, serializeStyles } from './serializer';
import { NumberingManager, serializeNumbering } from './numbering';
import { serializeSettings } from './settings';
import { serializeHeader, serializeFooter } from './header-footer';
import { serializeFootnotes, serializeEndnotes, NoteManager } from './footnotes';
import { WmlHyperlink } from './hyperlinks';
import { createCaptionParagraph, createTableOfContents } from './helpers';
import type { DocxSettings } from './settings';
import type { WmlAbstractNum } from './numbering';
import type { WmlHeaderFooterType, WmlHeaderFooterRef } from './header-footer';
import type { WmlFootnoteProperties, WmlEndnoteProperties, WmlFootnoteRef } from './footnotes';
import type { WmlHyperlink as WmlHyperlinkType } from './hyperlinks';
import type {
  WmlDocument, WmlBody, WmlBlockElement,
  WmlParagraph, WmlParagraphProperties, WmlParagraphContent,
  WmlRun, WmlRunProperties,
  WmlText, WmlBreak,
  WmlTable, WmlTableRow, WmlTableCell,
  WmlTableProperties, WmlTableGrid, WmlTableRowProperties, WmlTableCellProperties,
  WmlSectionProperties, WmlPageSize, WmlPageMargins,
  WmlStyles, WmlStyle, WmlStyleType, WmlDocDefaults, WmlTextDirectionValue,
} from './types';
import type { Twips, HalfPoints, HexColor, SignedTwips } from '../shared/types';
import { inToTwips, ptToTwips, ptToHalfPt, cmToTwips } from '../shared/types';

// Re-export measurement helpers for consumer convenience
export { inToTwips, ptToTwips, ptToHalfPt, cmToTwips };

// ---------------------------------------------------------------------------
// Core properties interface
// Source: opc-coreProperties.xsd — CT_CoreProperties (fully read)
// ---------------------------------------------------------------------------

export interface DocxCoreProperties {
  title?: string;
  subject?: string;
  creator?: string;
  description?: string;
  keywords?: string;
  category?: string;
  revision?: string;
  created?: Date;
  modified?: Date;
  lastModifiedBy?: string;
}

// ---------------------------------------------------------------------------
// Extended properties interface
// Source: shared-documentPropertiesExtended.xsd — CT_Properties (fully read)
// ---------------------------------------------------------------------------

export interface DocxExtendedProperties {
  application?: string;
  appVersion?: string;
  company?: string;
  manager?: string;
}

// ---------------------------------------------------------------------------
// Core properties serializer
// Source: opc-coreProperties.xsd (fully read)
// ---------------------------------------------------------------------------

function serializeCoreProperties(props: DocxCoreProperties): string {
  const children = [];

  if (props.title)          children.push(el('dc:title',          { 'xmlns:dc': NS.DC }, props.title));
  if (props.subject)        children.push(el('dc:subject',        { 'xmlns:dc': NS.DC }, props.subject));
  if (props.creator)        children.push(el('dc:creator',        { 'xmlns:dc': NS.DC }, props.creator));
  if (props.description)    children.push(el('dc:description',    { 'xmlns:dc': NS.DC }, props.description));
  if (props.keywords)       children.push(el('cp:keywords',       {}, props.keywords));
  if (props.category)       children.push(el('cp:category',       {}, props.category));
  if (props.revision)       children.push(el('cp:revision',       {}, props.revision));
  if (props.lastModifiedBy) children.push(el('cp:lastModifiedBy', {}, props.lastModifiedBy));

  if (props.created) {
    children.push(el('dcterms:created', {
      'xmlns:dcterms': NS.DCTERMS,
      'xmlns:xsi': NS.XSI,
      'xsi:type': 'dcterms:W3CDTF',
    }, props.created.toISOString()));
  }
  if (props.modified) {
    children.push(el('dcterms:modified', {
      'xmlns:dcterms': NS.DCTERMS,
      'xmlns:xsi': NS.XSI,
      'xsi:type': 'dcterms:W3CDTF',
    }, props.modified.toISOString()));
  }

  const root = el('cp:coreProperties', {
    'xmlns:cp':      NS.CORE_PROPS,
    'xmlns:dc':      NS.DC,
    'xmlns:dcterms': NS.DCTERMS,
    'xmlns:xsi':     NS.XSI,
  }, ...children);

  return xmlDocument(root);
}

// ---------------------------------------------------------------------------
// Extended properties serializer
// Source: shared-documentPropertiesExtended.xsd (fully read)
// ---------------------------------------------------------------------------

function serializeExtendedProperties(props: DocxExtendedProperties): string {
  const children = [];
  if (props.application) children.push(el('Application', {}, props.application));
  if (props.appVersion)  children.push(el('AppVersion',  {}, props.appVersion));
  if (props.company)     children.push(el('Company',     {}, props.company));
  if (props.manager)     children.push(el('Manager',     {}, props.manager));

  const root = el('Properties', { 'xmlns': NS.EXT_PROPS }, ...children);
  return xmlDocument(root);
}

// ---------------------------------------------------------------------------
// DocxBuilder
// ---------------------------------------------------------------------------

export class DocxBuilder {
  private blocks: WmlBlockElement[] = [];
  private styles: WmlStyle[] = [];
  private docDefaults?: WmlDocDefaults;
  private sectPr: WmlSectionProperties = {
    pgSz:  { w: 12240, h: 15840 },   // US Letter default
    pgMar: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
  };
  private coreProps: DocxCoreProperties = {};
  private extProps: DocxExtendedProperties = { application: 'ooxmlgen', appVersion: '1.0' };
  private numberingMgr = new NumberingManager();
  private settings: DocxSettings = {};
  private noteMgr = new NoteManager();
  /** Batch 2: header/footer parts */
  private hdrFtrParts: Array<{
    partPath: string;
    content: WmlBlockElement[];
    isHeader: boolean;
    type: WmlHeaderFooterType;
  }> = [];
  private hdrFtrCounter = 1;
  /** Batch 4: hyperlink relationships tracked for document.xml.rels */
  private hyperlinkRels: Array<{ url: string; placeholderId: string }> = [];
  private hyperlinkCounter = 1;
  // ---- Page setup ----

  /** Set page size. Use ptToTwips() or inToTwips() helpers. */
  setPageSize(pgSz: WmlPageSize): this {
    this.sectPr = { ...this.sectPr, pgSz };
    return this;
  }

  /** Set page margins. All values in twips. */
  setPageMargins(pgMar: WmlPageMargins): this {
    this.sectPr = { ...this.sectPr, pgMar };
    return this;
  }

  /** Set section type. */
  setSectionType(type: WmlSectionProperties['type']): this {
    this.sectPr = { ...this.sectPr, type };
    return this;
  }

  /** Set section text direction. */
  setSectionTextDirection(textDirection: WmlTextDirectionValue): this {
    this.sectPr = { ...this.sectPr, textDirection };
    return this;
  }

  // ---- Document properties ----

  setCoreProperties(props: DocxCoreProperties): this {
    this.coreProps = { ...this.coreProps, ...props };
    return this;
  }

  setExtendedProperties(props: DocxExtendedProperties): this {
    this.extProps = { ...this.extProps, ...props };
    return this;
  }

  // ---- Numbering (Batch 1) ----

  /**
   * Define a numbering definition and return its numId.
   * Use the returned numId in paragraph pPr.numPr.numId.
   * Source: CT_AbstractNum + CT_Num (wml.xsd lines 2989, 3004)
   *
   * @example
   * const numId = builder.defineNumbering({
   *   multiLevelType: 'singleLevel',
   *   levels: [{ ilvl: 0, numFmt: 'decimal', lvlText: '%1.', ... }]
   * });
   * builder.addBlock(paragraph(['Item 1'], { numPr: { numId, ilvl: 0 } }));
   */
  defineNumbering(def: Omit<WmlAbstractNum, 'abstractNumId'>): number {
    return this.numberingMgr.define(def);
  }

  // ---- Settings (Batch 1) ----

  /**
   * Configure document-level settings (word/settings.xml).
   * Source: CT_Settings (wml.xsd line 2733)
   *
   * @example
   * // For a document with TOC and page numbers:
   * builder.configureSettings({ updateFields: true });
   *
   * // For an Arabic RTL document:
   * builder.configureSettings({
   *   bidi: true,
   *   themeFontLang: { bidi: 'ar-SA' }
   * });
   */
  configureSettings(settings: DocxSettings): this {
    this.settings = { ...this.settings, ...settings };
    return this;
  }

  // ---- Headers and Footers (Batch 2) ----

  /**
   * Add a header part and return a WmlHeaderFooterRef to use in sectPr.
   * Source: CT_HdrFtr (wml.xsd line 1528), CT_HdrFtrRef (line 1514)
   *
   * @param type - 'default' (all pages), 'first' (first page), 'even' (even pages)
   * @param content - Block-level content for the header (paragraphs, tables)
   * @returns WmlHeaderFooterRef — add to sectPr.headerReference[]
   *
   * @example
   * const hdrRef = builder.addHeader('default', [
   *   paragraph([{ text: 'My Document', rPr: { b: true } }], { jc: 'center' })
   * ]);
   * builder.setSectionProperties({ ...sectPr, headerReference: [hdrRef] });
   */
  addHeader(type: WmlHeaderFooterType, content: WmlBlockElement[]): WmlHeaderFooterRef {
    const n = this.hdrFtrCounter++;
    const partPath = `/word/header${n}.xml`;
    this.hdrFtrParts.push({ partPath, content, isHeader: true, type });
    // rId will be assigned during build() — return a placeholder ref
    return { type, id: `__hdr_${n}__` };
  }

  /**
   * Add a footer part and return a WmlHeaderFooterRef to use in sectPr.
   * Source: CT_HdrFtr (wml.xsd line 1528), CT_HdrFtrRef (line 1514)
   *
   * @param type - 'default' (all pages), 'first' (first page), 'even' (even pages)
   * @param content - Block-level content for the footer (paragraphs, tables)
   * @returns WmlHeaderFooterRef — add to sectPr.footerReference[]
   *
   * @example
   * const ftrRef = builder.addFooter('default', [
   *   paragraph([...pageXofYRuns()], { jc: 'center' })
   * ]);
   * builder.setSectionProperties({ ...sectPr, footerReference: [ftrRef] });
   */
  addFooter(type: WmlHeaderFooterType, content: WmlBlockElement[]): WmlHeaderFooterRef {
    const n = this.hdrFtrCounter++;
    const partPath = `/word/footer${n}.xml`;
    this.hdrFtrParts.push({ partPath, content, isHeader: false, type });
    return { type, id: `__ftr_${n}__` };
  }

  /**
   * Convenience: set the final section properties (page size, margins, header/footer refs, etc.)
   * Replaces the current final sectPr entirely.
   */
  setSectionProperties(sectPr: WmlSectionProperties): this {
    this.sectPr = sectPr;
    return this;
  }

  // ---- Footnotes and Endnotes (Batch 3) ----

  /**
   * Add a footnote and return a run containing the reference mark.
   * Place the returned run inside a paragraph in the document body.
   * Source: CT_FtnEdn (wml.xsd line 2491), CT_FtnEdnRef (line 2484)
   *
   * @param content - Block-level content of the footnote (paragraphs, tables)
   * @returns A WmlRun containing w:footnoteReference — insert into a paragraph
   *
   * @example
   * const noteRun = builder.addFootnote([paragraph(['See Smith (2020).'])]);
   * builder.addBlock(paragraph(['Main text', noteRun]));
   */
  addFootnote(content: WmlBlockElement[]): WmlRun {
    const ref = this.noteMgr.addFootnote(content);
    return { _type: 'run', content: [ref] };
  }

  /**
   * Add an endnote and return a run containing the reference mark.
   * Source: CT_FtnEdn (wml.xsd line 2491), CT_FtnEdnRef (line 2484)
   *
   * @param content - Block-level content of the endnote
   * @returns A WmlRun containing w:endnoteReference — insert into a paragraph
   */
  addEndnote(content: WmlBlockElement[]): WmlRun {
    const ref = this.noteMgr.addEndnote(content);
    return { _type: 'run', content: [ref] };
  }

  /**
   * Configure footnote properties (position, number format, restart behavior).
   * Source: CT_FtnProps (wml.xsd line 2506)
   * These are emitted in word/settings.xml when set.
   */
  configureFootnotes(props: WmlFootnoteProperties): this {
    this.noteMgr.setFootnoteProperties(props);
    return this;
  }

  /**
   * Configure endnote properties.
   * Source: CT_EdnProps (wml.xsd line 2513)
   */
  configureEndnotes(props: WmlEndnoteProperties): this {
    this.noteMgr.setEndnoteProperties(props);
    return this;
  }

  // ---- Batch 4 helpers ----

  /**
   * Create a caption paragraph with a SEQ field prefix.
   * Reusable for figures, tables, equations, etc.
   */
  createCaptionParagraph(
    label: string,
    captionText: string,
    options: {
      pStyle?: string;
      pPr?: WmlParagraphProperties;
      rPr?: WmlRunProperties;
      seqOptions?: Parameters<typeof import('./fields').seqField>[1];
    } = {},
  ): WmlParagraph {
    return createCaptionParagraph(label, captionText, options, paragraph);
  }

  /**
   * Create a TOC placeholder paragraph.
   * The actual TOC is generated by Word when updateFields is true in settings.
   */
  createTableOfContents(options: {
    headingLevels?: number[];
    customSwitches?: string;
    pStyle?: string;
    pPr?: WmlParagraphProperties;
    rPr?: WmlRunProperties;
  } = {}): WmlParagraph {
    return createTableOfContents(options, paragraph);
  }

  // ---- Hyperlinks (Batch 4) ----

  /**
   * Register an external hyperlink relationship and return a placeholder rId.
   * The placeholder is resolved to the real rId during build().
   * Source: CT_Hyperlink (wml.xsd line 1218), CT_Rel (line 1196)
   *
   * @param url - The target URL (e.g. "https://example.com")
   * @returns A placeholder rId — assign to WmlHyperlink.rId
   *
   * @example
   * const rId = builder.addHyperlinkRelationship('https://example.com');
   * const link: WmlHyperlink = { _type: 'hyperlink', rId, runs: [...] };
   * builder.addBlock(paragraph([link]));
   */
  addHyperlinkRelationship(url: string): string {
    const n = this.hyperlinkCounter++;
    const placeholder = `__hlink_${n}__`;
    this.hyperlinkRels.push({ url, placeholderId: placeholder });
    return placeholder;
  }

  // ---- Batch 5: Document Defaults ----

  /**
   * Configure document-level defaults for runs and paragraphs.
   * Source: CT_DocDefaults (wml.xsd line 2687)
   *
   * @param docDefaults - Default properties applied to all content without explicit formatting
   * @returns this for chaining
   *
   * @example
   * // Arabic document with RTL defaults
   * builder.configureDocDefaults({
   *   rPrDefault: { rtl: true, rFonts: { cs: 'Arial Unicode MS' } },
   *   pPrDefault: { bidi: true, jc: 'right' }
   * });
   */
  configureDocDefaults(docDefaults: WmlDocDefaults): this {
    this.docDefaults = docDefaults;
    return this;
  }

  // ---- Styles ----

  /** Add a style definition. */
  addStyle(style: WmlStyle): this {
    this.styles.push(style);
    return this;
  }

  // ---- Content ----

  /** Add a pre-built block element (paragraph or table). */
  addBlock(block: WmlBlockElement): this {
    this.blocks.push(block);
    return this;
  }

  /**
   * Convenience: add a paragraph with optional style and runs.
   * For full control, build a WmlParagraph and call addBlock().
   */
  addParagraph(options: {
    style?: string;
    pPr?: WmlParagraphProperties;
    runs?: Array<{ text: string; rPr?: WmlRunProperties }>;
  } = {}): this {
    const pPr: WmlParagraphProperties | undefined = options.pPr
      ?? (options.style ? { pStyle: options.style } : undefined);

    const content: WmlParagraphContent[] = (options.runs ?? []).map(r => ({
      _type: 'run' as const,
      rPr: r.rPr && Object.values(r.rPr).some(v => v !== undefined) ? r.rPr : undefined,
      content: [{ _type: 'text' as const, text: r.text, space: 'preserve' as const }],
    }));

    this.blocks.push({ _type: 'paragraph', pPr, content });
    return this;
  }

  /**
   * Convenience: add a table.
   * rows: array of rows, each row is an array of cells.
   * Each cell is an array of WmlBlockElement (usually paragraphs).
   */
  addTable(options: {
    tblPr?: WmlTableProperties;
    colWidths: Twips[];
    rows: Array<{
      trPr?: WmlTableRowProperties;
      cells: Array<{
        tcPr?: WmlTableCellProperties;
        content: WmlBlockElement[];
      }>;
    }>;
  }): this {
    const tblGrid: WmlTableGrid = { cols: options.colWidths.map(w => ({ w })) };
    const tblPr: WmlTableProperties = options.tblPr ?? {};

    const rows: WmlTableRow[] = options.rows.map(r => ({
      _type: 'tableRow' as const,
      trPr: r.trPr,
      cells: r.cells.map(c => ({
        _type: 'tableCell' as const,
        tcPr: c.tcPr,
        // SPEC: cell must contain at least one block element
        content: c.content.length > 0 ? c.content : [emptyParagraph()],
      })),
    }));

    this.blocks.push({ _type: 'table', tblPr, tblGrid, rows });
    return this;
  }

  // ---- Build ----

  /** Build and return the DOCX as a Buffer. */
  async build(): Promise<Buffer> {
    const pkg = new OpcPackage();

    // Package-level relationships
    pkg.addPackageRelationship(REL_TYPES.OFFICE_DOCUMENT,    'word/document.xml');
    pkg.addPackageRelationship(REL_TYPES.CORE_PROPERTIES,    'docProps/core.xml');
    pkg.addPackageRelationship(REL_TYPES.EXTENDED_PROPERTIES, 'docProps/app.xml');

    // Batch 2: emit header/footer parts and build rId map
    // Map: placeholder id → real rId
    const rIdMap = new Map<string, string>();
    for (const hf of this.hdrFtrParts) {
      const xml = hf.isHeader
        ? serializeHeader(hf.content)
        : serializeFooter(hf.content);
      const contentType = hf.isHeader ? CONTENT_TYPES.HEADER : CONTENT_TYPES.FOOTER;
      const relType     = hf.isHeader ? REL_TYPES.HEADER      : REL_TYPES.FOOTER;
      const zipName     = hf.partPath.slice(1); // strip leading /
      const target      = zipName.replace('word/', ''); // relative to word/

      pkg.addPart(hf.partPath, contentType, xml);
      const rId = pkg.addPartRelationship('/word/document.xml', relType, target);

      // Extract placeholder key from partPath: /word/header1.xml → __hdr_1__
      const match = hf.partPath.match(/\/(header|footer)(\d+)\.xml$/);
      if (match) {
        const placeholder = `__${match[1] === 'header' ? 'hdr' : 'ftr'}_${match[2]}__`;
        rIdMap.set(placeholder, rId);
      }
    }

    // Batch 4: emit hyperlink relationships and build rId map for hyperlinks
    for (const hl of this.hyperlinkRels) {
      const rId = pkg.addPartRelationship('/word/document.xml', REL_TYPES.HYPERLINK, hl.url, 'External');
      rIdMap.set(hl.placeholderId, rId);
    }

    // Resolve placeholder rIds in sectPr (final section)
    const resolvedSectPr = this.resolveSectPrRefs(this.sectPr, rIdMap);

    // Also resolve in any mid-document section breaks (pPr.sectPr) and hyperlink rIds in paragraph content
    const resolvedBlocks = this.resolveBlockRefs(this.blocks, rIdMap);

    // word/document.xml
    const doc: WmlDocument = {
      conformance: 'strict',
      body: { content: resolvedBlocks, sectPr: resolvedSectPr },
    };
    pkg.addPart('/word/document.xml', CONTENT_TYPES.DOCX, serializeDocument(doc));

    // word/styles.xml
    const stylesXml = serializeStyles({ styles: this.styles, docDefaults: this.docDefaults });
    pkg.addPart('/word/styles.xml', CONTENT_TYPES.STYLES, stylesXml);
    pkg.addPartRelationship('/word/document.xml', REL_TYPES.STYLES, 'styles.xml');

    // word/numbering.xml — emitted only when numbering definitions exist
    if (this.numberingMgr.hasDefinitions()) {
      const numberingXml = serializeNumbering(this.numberingMgr.toWmlNumbering());
      pkg.addPart('/word/numbering.xml', CONTENT_TYPES.NUMBERING, numberingXml);
      pkg.addPartRelationship('/word/document.xml', REL_TYPES.NUMBERING, 'numbering.xml');
    }

    // word/settings.xml — emitted when any settings are configured
    const hasSettings = Object.values(this.settings).some(v => v !== undefined);
    if (hasSettings) {
      const settingsXml = serializeSettings(this.settings);
      pkg.addPart('/word/settings.xml', CONTENT_TYPES.SETTINGS, settingsXml);
      pkg.addPartRelationship('/word/document.xml', REL_TYPES.SETTINGS, 'settings.xml');
    }

    // word/footnotes.xml — emitted when footnotes exist (Batch 3)
    if (this.noteMgr.hasFootnotes()) {
      const footnotesXml = serializeFootnotes(
        this.noteMgr.getFootnotes(),
        this.noteMgr.getFootnoteProps(),
      );
      pkg.addPart('/word/footnotes.xml', CONTENT_TYPES.FOOTNOTES, footnotesXml);
      pkg.addPartRelationship('/word/document.xml', REL_TYPES.FOOTNOTES, 'footnotes.xml');
    }

    // word/endnotes.xml — emitted when endnotes exist (Batch 3)
    if (this.noteMgr.hasEndnotes()) {
      const endnotesXml = serializeEndnotes(
        this.noteMgr.getEndnotes(),
        this.noteMgr.getEndnoteProps(),
      );
      pkg.addPart('/word/endnotes.xml', CONTENT_TYPES.ENDNOTES, endnotesXml);
      pkg.addPartRelationship('/word/document.xml', REL_TYPES.ENDNOTES, 'endnotes.xml');
    }

    // docProps/core.xml
    pkg.addPart('/docProps/core.xml', CONTENT_TYPES.CORE_PROPS, serializeCoreProperties(this.coreProps));

    // docProps/app.xml
    pkg.addPart('/docProps/app.xml', CONTENT_TYPES.EXT_PROPS, serializeExtendedProperties(this.extProps));

    return pkg.build();
  }

  /** Resolve placeholder rIds in a sectPr's headerReference/footerReference arrays. */
  private resolveSectPrRefs(
    sectPr: WmlSectionProperties,
    rIdMap: Map<string, string>,
  ): WmlSectionProperties {
    if (!sectPr.headerReference && !sectPr.footerReference) return sectPr;
    return {
      ...sectPr,
      headerReference: sectPr.headerReference?.map(ref => ({
        ...ref,
        id: rIdMap.get(ref.id) ?? ref.id,
      })),
      footerReference: sectPr.footerReference?.map(ref => ({
        ...ref,
        id: rIdMap.get(ref.id) ?? ref.id,
      })),
    };
  }

  /** Resolve placeholder rIds in pPr.sectPr for mid-document section breaks and hyperlink rIds in paragraph content. */
  private resolveBlockRefs(
    blocks: WmlBlockElement[],
    rIdMap: Map<string, string>,
  ): WmlBlockElement[] {
    if (rIdMap.size === 0) return blocks;
    return blocks.map(block => {
      if (block._type !== 'paragraph') return block;
      const resolvedContent = block.content.map(c => {
        if (c._type === 'run') return c;
        if (c._type === 'hyperlink') {
          return {
            ...c,
            rId: c.rId ? (rIdMap.get(c.rId) ?? c.rId) : c.rId,
          } as WmlHyperlinkType;
        }
        return c;
      });
      const hasSectPr = block.pPr?.sectPr;
      return {
        ...block,
        content: resolvedContent,
        pPr: hasSectPr
          ? { ...block.pPr, sectPr: this.resolveSectPrRefs(block.pPr!.sectPr!, rIdMap) }
          : block.pPr,
      };
    });
  }
}

// ---------------------------------------------------------------------------
// Factory helpers — convenience constructors for common WML constructs
// ---------------------------------------------------------------------------

/** Create an empty paragraph (required as placeholder in table cells etc.) */
export function emptyParagraph(pPr?: WmlParagraphProperties): WmlParagraph {
  return { _type: 'paragraph', pPr, content: [] };
}

/** Create a text run */
export function textRun(text: string, rPr?: WmlRunProperties): WmlRun {
  // Guard: don't attach rPr if it has no actual properties set
  const effectiveRPr = rPr && Object.values(rPr).some(v => v !== undefined) ? rPr : undefined;
  return {
    _type: 'run',
    rPr: effectiveRPr,
    content: [{ _type: 'text', text, space: 'preserve' }],
  };
}

/** Create a paragraph with text runs, hyperlinks, or pre-built objects */
export function paragraph(
  runs: Array<string | { text: string; rPr?: WmlRunProperties } | WmlRun | import('./hyperlinks').WmlHyperlink>,
  pPr?: WmlParagraphProperties,
): WmlParagraph {
  const content: WmlParagraphContent[] = runs.map(r => {
    if (typeof r === 'string') return textRun(r);
    if ('_type' in r) {
      if (r._type === 'run') return r as WmlRun;
      if (r._type === 'hyperlink') return r as import('./hyperlinks').WmlHyperlink;
    }
    return textRun((r as { text: string; rPr?: WmlRunProperties }).text,
                   (r as { text: string; rPr?: WmlRunProperties }).rPr);
  });
  return { _type: 'paragraph', pPr, content };
}

/** Create a page break paragraph */
export function pageBreak(): WmlParagraph {
  const br: WmlBreak = { _type: 'break', breakType: 'page' };
  return {
    _type: 'paragraph',
    content: [{ _type: 'run', content: [br] }],
  };
}

// Re-export types for consumer convenience
export type {
  WmlDocument, WmlBody, WmlBlockElement,
  WmlParagraph, WmlParagraphProperties,
  WmlRun, WmlRunProperties,
  WmlText, WmlBreak,
  WmlTable, WmlTableRow, WmlTableCell,
  WmlTableProperties, WmlTableGrid, WmlTableRowProperties, WmlTableCellProperties,
  WmlSectionProperties, WmlPageSize, WmlPageMargins,
  WmlStyles, WmlStyle, WmlStyleType,
  Twips, HalfPoints, HexColor, SignedTwips,
};