/**
 * Minimal XML builder for OOXML serialization.
 *
 * Traceability: SPEC_TRACEABILITY.md — XmlBuilder
 * Source justification: All OOXML parts (wml.xsd, opc-*.xsd) are XML documents
 * requiring well-formed output with correct namespace declarations.
 *
 * Design constraints derived from spec:
 * - xml:space="preserve" must be expressible (CT_Text in wml.xsd)
 * - Namespace prefixes must be stable (w:, r:, wp:, a:, etc.)
 * - Attribute order is not normative but must be consistent
 * - No external dependencies — pure Node.js
 */

/** A single XML attribute: name → value */
export type XmlAttrs = Record<string, string | number | boolean | undefined>;

/** An XML node: either an element or a text node */
export type XmlNode = XmlElement | string;

export interface XmlElement {
  tag: string;
  attrs: XmlAttrs;
  children: XmlNode[];
}

/** Create an XML element node */
export function el(tag: string, attrs: XmlAttrs = {}, ...children: XmlNode[]): XmlElement {
  return { tag, attrs, children };
}

/** Escape XML special characters in text content */
function escapeText(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Escape XML special characters in attribute values */
function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Serialize an XmlNode tree to a string */
export function serialize(node: XmlNode, indent = false, depth = 0): string {
  if (typeof node === 'string') {
    return escapeText(node);
  }

  const { tag, attrs, children } = node;
  const pad = indent ? '  '.repeat(depth) : '';
  const nl = indent ? '\n' : '';

  // Build attribute string — skip undefined values
  const attrStr = Object.entries(attrs)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${escapeAttr(String(v))}"`)
    .join('');

  if (children.length === 0) {
    return `${pad}<${tag}${attrStr}/>`;
  }

  // Single text child — inline
  if (children.length === 1 && typeof children[0] === 'string') {
    return `${pad}<${tag}${attrStr}>${escapeText(children[0])}</${tag}>`;
  }

  const inner = children
    .map(c => serialize(c, indent, depth + 1))
    .join(nl);

  return `${pad}<${tag}${attrStr}>${nl}${inner}${nl}${pad}</${tag}>`;
}

/** Produce a complete XML document string with declaration */
export function xmlDocument(root: XmlElement, indent = false): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serialize(root, indent)}`;
}
