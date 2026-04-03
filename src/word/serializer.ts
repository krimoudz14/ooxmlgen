/**
 * WML Serializer — converts the WML object model to XML strings.
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd (partially read — see repo-manifest.json)
 * Traceability: SPEC_TRACEABILITY.md — WmlSerializer
 *
 * V1 BOUNDARY: Serializes only constructs within the confirmed v1 scope.
 * Deferred features are explicitly noted and will throw or be silently skipped.
 *
 * Key spec constraints enforced here:
 * - w:t xml:space="preserve" when text has leading/trailing whitespace (CT_Text)
 * - w:tbl requires w:tblPr and w:tblGrid as first children (CT_Tbl)
 * - w:tc must contain at least one block element (CT_Tc)
 * - w:sectPr must be last child of w:body (CT_Body)
 * - Namespace prefix w: on all WML elements
 */

import { el, xmlDocument, XmlElement, XmlAttrs } from '../shared/xml';
import { NS } from '../shared/namespaces';
import type {
  WmlDocument, WmlBody, WmlBlockElement,
  WmlParagraph, WmlParagraphProperties,
  WmlRun, WmlRunProperties, WmlRunContent,
  WmlText, WmlBreak,
  WmlTable, WmlTableRow, WmlTableCell,
  WmlTableProperties, WmlTableGrid, WmlTableRowProperties, WmlTableCellProperties,
  WmlSectionProperties, WmlPageSize, WmlPageMargins,
  WmlStyles, WmlStyle, WmlDocDefaults,
  WmlColor, WmlFonts, WmlSpacing, WmlIndentation,
  WmlBorder, WmlShading, WmlTableBorders, WmlTableWidth,
  WmlTableCellBorders, WmlTableCellMargins,
  WmlParagraphContent,
} from './types';
import type { WmlHyperlink } from './hyperlinks';
// Shorthand: create a w:-prefixed element
export function w(localName: string, attrs: XmlAttrs = {}, ...children: (XmlElement | string)[]): XmlElement {
  return el(`w:${localName}`, attrs, ...children);
}

// Serialize a boolean CT_OnOff element — omit if undefined, omit val attr if true
export function onOff(name: string, val: boolean | undefined): XmlElement | null {
  if (val === undefined) return null;
  return val ? w(name) : w(name, { 'w:val': '0' });
}

// Serialize a CT_String element
export function strElem(name: string, val: string | undefined): XmlElement | null {
  if (val === undefined) return null;
  return w(name, { 'w:val': val });
}

// Serialize a CT_DecimalNumber element
function decNum(name: string, val: number | undefined): XmlElement | null {
  if (val === undefined) return null;
  return w(name, { 'w:val': String(val) });
}

// Filter nulls from children array
function compact(items: (XmlElement | null | undefined)[]): XmlElement[] {
  return items.filter((x): x is XmlElement => x != null);
}

// ---------------------------------------------------------------------------
// Shared sub-type serializers
// ---------------------------------------------------------------------------

function serializeColor(color: WmlColor): XmlElement {
  return w('color', { 'w:val': color.val });
}

function serializeFonts(fonts: WmlFonts): XmlElement {
  const attrs: XmlAttrs = {};
  if (fonts.ascii)    attrs['w:ascii']    = fonts.ascii;
  if (fonts.hAnsi)    attrs['w:hAnsi']    = fonts.hAnsi;
  if (fonts.eastAsia) attrs['w:eastAsia'] = fonts.eastAsia;
  if (fonts.cs)       attrs['w:cs']       = fonts.cs;
  return w('rFonts', attrs);
}

function serializeSpacing(spacing: WmlSpacing): XmlElement {
  const attrs: XmlAttrs = {};
  if (spacing.before !== undefined)           attrs['w:before']           = spacing.before;
  if (spacing.after !== undefined)            attrs['w:after']            = spacing.after;
  if (spacing.line !== undefined)             attrs['w:line']             = spacing.line;
  if (spacing.lineRule !== undefined)         attrs['w:lineRule']         = spacing.lineRule;
  if (spacing.beforeAutospacing !== undefined) attrs['w:beforeAutospacing'] = spacing.beforeAutospacing ? '1' : '0';
  if (spacing.afterAutospacing !== undefined)  attrs['w:afterAutospacing']  = spacing.afterAutospacing  ? '1' : '0';
  return w('spacing', attrs);
}

function serializeIndentation(ind: WmlIndentation): XmlElement {
  const attrs: XmlAttrs = {};
  if (ind.start !== undefined)     attrs['w:start']     = ind.start;
  if (ind.end !== undefined)       attrs['w:end']        = ind.end;
  if (ind.hanging !== undefined)   attrs['w:hanging']   = ind.hanging;
  if (ind.firstLine !== undefined) attrs['w:firstLine'] = ind.firstLine;
  return w('ind', attrs);
}

function serializeBorder(name: string, border: WmlBorder): XmlElement {
  const attrs: XmlAttrs = { 'w:val': border.val };
  if (border.color !== undefined) attrs['w:color'] = border.color;
  if (border.sz !== undefined)    attrs['w:sz']    = border.sz;
  if (border.space !== undefined) attrs['w:space'] = border.space;
  if (border.shadow !== undefined) attrs['w:shadow'] = border.shadow ? '1' : '0';
  return w(name, attrs);
}

function serializeShading(shd: WmlShading): XmlElement {
  const attrs: XmlAttrs = { 'w:val': shd.val };
  if (shd.color !== undefined) attrs['w:color'] = shd.color;
  if (shd.fill !== undefined)  attrs['w:fill']  = shd.fill;
  return w('shd', attrs);
}

function serializeTableWidth(name: string, tw: WmlTableWidth): XmlElement {
  return w(name, { 'w:w': tw.w, 'w:type': tw.type });
}

// ---------------------------------------------------------------------------
// Run properties serializer
// Source: CT_RPr / EG_RPrBase (wml.xsd lines 1742-1790)
// ---------------------------------------------------------------------------

/** @internal exported for use by numbering serializer */
export function serializeRPr(rPr: WmlRunProperties): XmlElement {
  const children = compact([
    strElem('rStyle', rPr.rStyle),
    rPr.rFonts ? serializeFonts(rPr.rFonts) : null,
    onOff('b', rPr.b),
    onOff('bCs', rPr.bCs),
    onOff('i', rPr.i),
    onOff('iCs', rPr.iCs),
    onOff('caps', rPr.caps),
    onOff('smallCaps', rPr.smallCaps),
    onOff('strike', rPr.strike),
    onOff('dstrike', rPr.dstrike),
    onOff('outline', rPr.outline),
    onOff('shadow', rPr.shadow),
    onOff('emboss', rPr.emboss),
    onOff('imprint', rPr.imprint),
    onOff('noProof', rPr.noProof),
    onOff('snapToGrid', rPr.snapToGrid),
    onOff('vanish', rPr.vanish),
    onOff('webHidden', rPr.webHidden),
    rPr.color ? serializeColor(rPr.color) : null,
    rPr.spacing !== undefined ? w('spacing', { 'w:val': rPr.spacing }) : null,
    rPr.kern !== undefined ? w('kern', { 'w:val': rPr.kern }) : null,
    rPr.position !== undefined ? w('position', { 'w:val': rPr.position }) : null,
    rPr.sz !== undefined ? w('sz', { 'w:val': rPr.sz }) : null,
    rPr.szCs !== undefined ? w('szCs', { 'w:val': rPr.szCs }) : null,
    rPr.highlight ? w('highlight', { 'w:val': rPr.highlight }) : null,
    rPr.u ? w('u', { 'w:val': rPr.u }) : null,
    rPr.shd ? serializeShading(rPr.shd) : null,
    rPr.bdr ? serializeBorder('bdr', rPr.bdr) : null,
    rPr.vertAlign ? w('vertAlign', { 'w:val': rPr.vertAlign }) : null,
    onOff('rtl', rPr.rtl),
    onOff('cs', rPr.cs),
    rPr.lang ? w('lang', {
      ...(rPr.lang.val      ? { 'w:val':      rPr.lang.val }      : {}),
      ...(rPr.lang.eastAsia ? { 'w:eastAsia': rPr.lang.eastAsia } : {}),
      ...(rPr.lang.bidi     ? { 'w:bidi':     rPr.lang.bidi }     : {}),
    }) : null,
  ]);
  return w('rPr', {}, ...children);
}

// ---------------------------------------------------------------------------
// Run content serializer
// Source: EG_RunInnerContent (wml.xsd lines 1667-1700)
// ---------------------------------------------------------------------------

function serializeRunContent(content: WmlRunContent): XmlElement {
  switch (content._type) {
    case 'text': {
      const t = content as WmlText;
      // SPEC: xml:space="preserve" required when text has leading/trailing whitespace
      const needsPreserve = t.space === 'preserve' || /^\s|\s$/.test(t.text);
      const attrs: XmlAttrs = needsPreserve ? { 'xml:space': 'preserve' } : {};
      return w('t', attrs, t.text);
    }
    case 'break': {
      const br = content as WmlBreak;
      const attrs: XmlAttrs = {};
      if (br.breakType) attrs['w:type'] = br.breakType;
      if (br.clear)     attrs['w:clear'] = br.clear;
      return w('br', attrs);
    }
    case 'tab':
      return w('tab');
    case 'cr':
      return w('cr');
    // Batch 2 additions — field infrastructure
    case 'fldChar': {
      const fc = content as import('./fields').WmlFldChar;
      const attrs: XmlAttrs = { 'w:fldCharType': fc.fldCharType };
      if (fc.fldLock) attrs['w:fldLock'] = '1';
      if (fc.dirty)   attrs['w:dirty']   = '1';
      return w('fldChar', attrs);
    }
    case 'instrText': {
      const it = content as import('./fields').WmlInstrText;
      // SPEC: xml:space="preserve" always set — field instructions contain spaces
      return w('instrText', { 'xml:space': 'preserve' }, it.text);
    }
    // Batch 3 additions — footnote/endnote references
    case 'footnoteReference': {
      const ref = content as import('./footnotes').WmlFootnoteRef;
      const attrs: XmlAttrs = { 'w:id': ref.id };
      if (ref.customMarkFollows) attrs['w:customMarkFollows'] = '1';
      return w('footnoteReference', attrs);
    }
    case 'endnoteReference': {
      const ref = content as import('./footnotes').WmlFootnoteRef;
      const attrs: XmlAttrs = { 'w:id': ref.id };
      if (ref.customMarkFollows) attrs['w:customMarkFollows'] = '1';
      return w('endnoteReference', attrs);
    }
  }
}

// ---------------------------------------------------------------------------
// Run serializer
// Source: CT_R (wml.xsd line 1703)
// ---------------------------------------------------------------------------

/** Returns true if rPr has at least one property set — avoids emitting empty w:rPr/ */
function rPrHasContent(rPr: WmlRunProperties): boolean {
  return Object.values(rPr).some(v => v !== undefined);
}

function serializeRun(run: WmlRun): XmlElement {
  const children: XmlElement[] = [];
  if (run.rPr && rPrHasContent(run.rPr)) children.push(serializeRPr(run.rPr));
  for (const c of run.content) {
    children.push(serializeRunContent(c));
  }
  return w('r', {}, ...children);
}

// ---------------------------------------------------------------------------
// Paragraph properties serializer
// Source: CT_PPrBase (wml.xsd line 1049), CT_PPr (wml.xsd line 1038)
// ---------------------------------------------------------------------------

function serializePPr(pPr: WmlParagraphProperties): XmlElement {
  const children = compact([
    strElem('pStyle', pPr.pStyle),
    onOff('keepNext', pPr.keepNext),
    onOff('keepLines', pPr.keepLines),
    onOff('pageBreakBefore', pPr.pageBreakBefore),
    onOff('widowControl', pPr.widowControl),
    onOff('suppressLineNumbers', pPr.suppressLineNumbers),
    // w:numPr — now serialized (Batch 1: numbering support)
    pPr.numPr ? w('numPr', {},
      w('ilvl', { 'w:val': pPr.numPr.ilvl }),
      w('numId', { 'w:val': pPr.numPr.numId }),
    ) : null,
    pPr.shd ? serializeShading(pPr.shd) : null,
    onOff('suppressAutoHyphens', pPr.suppressAutoHyphens),
    onOff('bidi', pPr.bidi),
    onOff('contextualSpacing', pPr.contextualSpacing),
    pPr.spacing ? serializeSpacing(pPr.spacing) : null,
    pPr.ind ? serializeIndentation(pPr.ind) : null,
    pPr.jc ? w('jc', { 'w:val': pPr.jc }) : null,
    pPr.outlineLvl !== undefined ? decNum('outlineLvl', pPr.outlineLvl) : null,
    // textDirection — Batch 1 addition
    pPr.textDirection ? w('textDirection', { 'w:val': pPr.textDirection }) : null,
  ]);
  // sectPr must be last child of pPr (CT_PPr constraint)
  const result = w('pPr', {}, ...children);
  if (pPr.sectPr) {
    result.children.push(serializeSectPr(pPr.sectPr));
  }
  return result;
}

/**
 * Serialize paragraph properties for use in numbering levels and style definitions
 * (CT_PPrGeneral — same as CT_PPrBase, no sectPr).
 * @internal exported for use by numbering serializer
 */
export function serializePPrGeneral(pPr: WmlParagraphProperties): XmlElement {
  return serializePPr(pPr);
}

// ---------------------------------------------------------------------------
// Paragraph serializer
// Source: CT_P (wml.xsd line 2159)
// ---------------------------------------------------------------------------

function serializeParagraph(p: WmlParagraph): XmlElement {
  const children: XmlElement[] = [];
  if (p.pPr) children.push(serializePPr(p.pPr));
  for (const c of p.content) {
    children.push(serializeParagraphContent(c));
  }
  return w('p', {}, ...children);
}

// ---------------------------------------------------------------------------
// Paragraph content serializer (runs + hyperlinks)
// Source: EG_PContent (wml.xsd ~line 2150)
// ---------------------------------------------------------------------------

function serializeParagraphContent(content: WmlParagraphContent): XmlElement {
  if (content._type === 'run') return serializeRun(content);
  if (content._type === 'hyperlink') return serializeHyperlink(content);
  // Exhaustiveness check — should never reach here
  throw new Error(`Unknown paragraph content type: ${(content as any)._type}`);
}

function serializeHyperlink(link: WmlHyperlink): XmlElement {
  const attrs: XmlAttrs = {};
  if (link.rId)    attrs['r:id']    = link.rId;
  if (link.anchor) attrs['w:anchor'] = link.anchor;
  if (link.tooltip) attrs['w:tooltip'] = link.tooltip;
  if (link.history !== undefined) attrs['w:history'] = link.history ? '1' : '0';

  const children = link.runs.map(serializeRun);
  return el('w:hyperlink', attrs, ...children);
}

// ---------------------------------------------------------------------------
// Table serializers
// Source: CT_Tbl (wml.xsd line 2434), CT_Row, CT_Tc
// ---------------------------------------------------------------------------

function serializeTableBorders(borders: WmlTableBorders): XmlElement {
  const children = compact([
    borders.top     ? serializeBorder('top',     borders.top)     : null,
    borders.start   ? serializeBorder('start',   borders.start)   : null,
    borders.bottom  ? serializeBorder('bottom',  borders.bottom)  : null,
    borders.end     ? serializeBorder('end',     borders.end)     : null,
    borders.insideH ? serializeBorder('insideH', borders.insideH) : null,
    borders.insideV ? serializeBorder('insideV', borders.insideV) : null,
  ]);
  return w('tblBorders', {}, ...children);
}

/** Serialize CT_TcBorders — Batch 1 addition */
function serializeTcBorders(borders: WmlTableCellBorders): XmlElement {
  const children = compact([
    borders.top     ? serializeBorder('top',     borders.top)     : null,
    borders.start   ? serializeBorder('start',   borders.start)   : null,
    borders.bottom  ? serializeBorder('bottom',  borders.bottom)  : null,
    borders.end     ? serializeBorder('end',     borders.end)     : null,
    borders.insideH ? serializeBorder('insideH', borders.insideH) : null,
    borders.insideV ? serializeBorder('insideV', borders.insideV) : null,
    borders.tl2br   ? serializeBorder('tl2br',   borders.tl2br)   : null,
    borders.tr2bl   ? serializeBorder('tr2bl',   borders.tr2bl)   : null,
  ]);
  return w('tcBorders', {}, ...children);
}

/** Serialize CT_TcMar — Batch 1 addition */
function serializeTcMar(mar: WmlTableCellMargins): XmlElement {
  const children = compact([
    mar.top    ? serializeTableWidth('top',    mar.top)    : null,
    mar.start  ? serializeTableWidth('start',  mar.start)  : null,
    mar.bottom ? serializeTableWidth('bottom', mar.bottom) : null,
    mar.end    ? serializeTableWidth('end',    mar.end)    : null,
  ]);
  return w('tcMar', {}, ...children);
}

function serializeTblPr(tblPr: WmlTableProperties): XmlElement {
  const children = compact([
    strElem('tblStyle', tblPr.tblStyle),
    tblPr.tblW ? serializeTableWidth('tblW', tblPr.tblW) : null,
    tblPr.jc ? w('jc', { 'w:val': tblPr.jc }) : null,
    tblPr.tblBorders ? serializeTableBorders(tblPr.tblBorders) : null,
    tblPr.shd ? serializeShading(tblPr.shd) : null,
    tblPr.tblLayout ? w('tblLayout', { 'w:type': tblPr.tblLayout }) : null,
  ]);
  return w('tblPr', {}, ...children);
}

function serializeTblGrid(grid: WmlTableGrid): XmlElement {
  const cols = grid.cols.map(c => w('gridCol', { 'w:w': c.w }));
  return w('tblGrid', {}, ...cols);
}

function serializeTrPr(trPr: WmlTableRowProperties): XmlElement {
  const children = compact([
    onOff('cantSplit', trPr.cantSplit),
    onOff('tblHeader', trPr.tblHeader),
    trPr.trHeight ? w('trHeight', {
      'w:val': trPr.trHeight.val,
      ...(trPr.trHeight.hRule ? { 'w:hRule': trPr.trHeight.hRule } : {}),
    }) : null,
    trPr.jc ? w('jc', { 'w:val': trPr.jc }) : null,
  ]);
  return w('trPr', {}, ...children);
}

function serializeTcPr(tcPr: WmlTableCellProperties): XmlElement {
  const children = compact([
    tcPr.tcW ? serializeTableWidth('tcW', tcPr.tcW) : null,
    tcPr.gridSpan !== undefined ? decNum('gridSpan', tcPr.gridSpan) : null,
    // vMerge — Batch 1 addition
    tcPr.vMerge !== undefined ? (
      tcPr.vMerge === 'restart'
        ? w('vMerge', { 'w:val': 'restart' })
        : w('vMerge', {})  // 'continue' = empty element
    ) : null,
    // tcBorders — Batch 1 addition
    tcPr.tcBorders ? serializeTcBorders(tcPr.tcBorders) : null,
    tcPr.shd ? serializeShading(tcPr.shd) : null,
    // noWrap — Batch 1 addition
    onOff('noWrap', tcPr.noWrap),
    // tcMar — Batch 1 addition
    tcPr.tcMar ? serializeTcMar(tcPr.tcMar) : null,
    // textDirection — Batch 1 addition
    tcPr.textDirection ? w('textDirection', { 'w:val': tcPr.textDirection }) : null,
    tcPr.vAlign ? w('vAlign', { 'w:val': tcPr.vAlign }) : null,
    // hideMark — Batch 1 addition
    onOff('hideMark', tcPr.hideMark),
  ]);
  return w('tcPr', {}, ...children);
}

function serializeTableCell(tc: WmlTableCell): XmlElement {
  const children: XmlElement[] = [];
  if (tc.tcPr) children.push(serializeTcPr(tc.tcPr));
  // SPEC: cell must contain at least one block element
  for (const block of tc.content) {
    children.push(serializeBlock(block));
  }
  return w('tc', {}, ...children);
}

function serializeTableRow(tr: WmlTableRow): XmlElement {
  const children: XmlElement[] = [];
  if (tr.trPr) children.push(serializeTrPr(tr.trPr));
  for (const tc of tr.cells) {
    children.push(serializeTableCell(tc));
  }
  return w('tr', {}, ...children);
}

function serializeTable(tbl: WmlTable): XmlElement {
  // SPEC: tblPr and tblGrid are required first children
  const children: XmlElement[] = [
    serializeTblPr(tbl.tblPr),
    serializeTblGrid(tbl.tblGrid),
    ...tbl.rows.map(serializeTableRow),
  ];
  return w('tbl', {}, ...children);
}

// ---------------------------------------------------------------------------
// Block element dispatcher
// ---------------------------------------------------------------------------

/** @internal exported for use by header/footer serializer */
export function serializeBlock(block: WmlBlockElement): XmlElement {
  switch (block._type) {
    case 'paragraph': return serializeParagraph(block);
    case 'table':     return serializeTable(block);
  }
}

// ---------------------------------------------------------------------------
// Section properties serializer
// Source: CT_SectPr (wml.xsd line 1566), EG_SectPrContents
// ---------------------------------------------------------------------------

function serializePageSz(pgSz: WmlPageSize): XmlElement {
  const attrs: XmlAttrs = {};
  if (pgSz.w !== undefined)      attrs['w:w']      = pgSz.w;
  if (pgSz.h !== undefined)      attrs['w:h']      = pgSz.h;
  if (pgSz.orient !== undefined) attrs['w:orient'] = pgSz.orient;
  return w('pgSz', attrs);
}

function serializePageMar(pgMar: WmlPageMargins): XmlElement {
  return w('pgMar', {
    'w:top':    pgMar.top,
    'w:right':  pgMar.right,
    'w:bottom': pgMar.bottom,
    'w:left':   pgMar.left,
    'w:header': pgMar.header,
    'w:footer': pgMar.footer,
    'w:gutter': pgMar.gutter,
  });
}

function serializeSectPr(sectPr: WmlSectionProperties): XmlElement {
  // headerReference / footerReference must come BEFORE EG_SectPrContents per CT_SectPr schema
  const hdrRefs = (sectPr.headerReference ?? []).map(ref =>
    w('headerReference', { 'w:type': ref.type, 'r:id': ref.id }),
  );
  const ftrRefs = (sectPr.footerReference ?? []).map(ref =>
    w('footerReference', { 'w:type': ref.type, 'r:id': ref.id }),
  );

  const contents = compact([
    sectPr.type ? w('type', { 'w:val': sectPr.type }) : null,
    sectPr.pgSz ? serializePageSz(sectPr.pgSz) : null,
    sectPr.pgMar ? serializePageMar(sectPr.pgMar) : null,
    sectPr.pgNumType ? w('pgNumType', {
      ...(sectPr.pgNumType.fmt   ? { 'w:fmt':   sectPr.pgNumType.fmt }           : {}),
      ...(sectPr.pgNumType.start !== undefined ? { 'w:start': sectPr.pgNumType.start } : {}),
    }) : null,
    onOff('titlePg', sectPr.titlePg),
    onOff('bidi', sectPr.bidi),
    sectPr.vAlign ? w('vAlign', { 'w:val': sectPr.vAlign }) : null,
    sectPr.textDirection ? w('textDirection', { 'w:val': sectPr.textDirection }) : null,
    onOff('rtlGutter', sectPr.rtlGutter),
  ]);

  return w('sectPr', {}, ...hdrRefs, ...ftrRefs, ...contents);
}

// ---------------------------------------------------------------------------
// Body serializer
// Source: CT_Body (wml.xsd line 3233)
// ---------------------------------------------------------------------------

function serializeBody(body: WmlBody): XmlElement {
  const children: XmlElement[] = body.content.map(serializeBlock);
  // SPEC: sectPr must be last child of body
  if (body.sectPr) children.push(serializeSectPr(body.sectPr));
  return w('body', {}, ...children);
}

// ---------------------------------------------------------------------------
// Document serializer
// Source: CT_Document (wml.xsd line 3439)
// ---------------------------------------------------------------------------

/**
 * Serialize a WmlDocument to a complete word/document.xml string.
 * Namespace declarations follow the pattern used by Word itself.
 */
export function serializeDocument(doc: WmlDocument): string {
  const attrs: XmlAttrs = {
    'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'xmlns:mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:o':   'urn:schemas-microsoft-com:office:office',
    'xmlns:r':   NS.R,
    'xmlns:m':   NS.M,
    'xmlns:v':   'urn:schemas-microsoft-com:vml',
    'xmlns:wp':  NS.WP,
    'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'xmlns:w':   NS.W,
    'mc:Ignorable': 'w14',
  };
  // NOTE: w:conformance="strict" is intentionally omitted — Word rejects it
  // even though it is valid per ECMA-376. Omitting it defaults to transitional.

  const root = el('w:document', attrs, serializeBody(doc.body));
  return xmlDocument(root);
}

// ---------------------------------------------------------------------------
// Styles serializer
// Source: CT_Styles (wml.xsd ~line 3110), CT_Style (wml.xsd line 3058)
// ---------------------------------------------------------------------------

function serializeStyle(style: WmlStyle): XmlElement {
  const attrs: XmlAttrs = {
    'w:type':    style.type,
    'w:styleId': style.styleId,
  };
  if (style.default) attrs['w:default'] = '1';

  const children = compact([
    w('name', { 'w:val': style.name }),
    strElem('basedOn', style.basedOn),
    strElem('next', style.next),
    strElem('link', style.link),
    style.pPr ? serializePPr(style.pPr) : null,
    style.rPr ? serializeRPr(style.rPr) : null,
    // tblPr / trPr / tcPr: DEFERRED_V1_1
  ]);

  return w('style', attrs, ...children);
}

/**
 * Serialize WmlDocDefaults to w:docDefaults element.
 */
function serializeDocDefaults(docDefaults: WmlDocDefaults): XmlElement {
  const children: XmlElement[] = [];
  if (docDefaults.rPrDefault) {
    children.push(el('w:rPrDefault', {}, serializeRPr(docDefaults.rPrDefault)));
  }
  if (docDefaults.pPrDefault) {
    children.push(el('w:pPrDefault', {}, serializePPrGeneral(docDefaults.pPrDefault)));
  }
  return el('w:docDefaults', {}, ...children);
}

/**
 * Serialize a WmlStyles to a complete word/styles.xml string.
 */
export function serializeStyles(styles: WmlStyles): string {
  const children = compact([
    styles.docDefaults ? serializeDocDefaults(styles.docDefaults) : null,
    ...styles.styles.map(serializeStyle),
  ]);
  const root = el('w:styles', {
    'xmlns:w': NS.W,
    'xmlns:r': NS.R,
  }, ...children);
  return xmlDocument(root);
}
