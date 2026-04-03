/**
 * OPC Content Types manager.
 *
 * Traceability: SPEC_TRACEABILITY.md — ContentTypeManager
 * Source: 5TH EDITION PART2/opc-contentTypes.xsd
 *   CT_Types: root element [Content_Types].xml
 *   CT_Default: Extension (required) + ContentType (required)
 *   CT_Override: PartName (required) + ContentType (required)
 *
 * Constraint from spec: every part MUST have a content type via Default or Override.
 */

import { el, xmlDocument, XmlElement } from '../shared/xml';
import { NS } from '../shared/namespaces';

export class ContentTypeManager {
  private defaults = new Map<string, string>(); // ext → contentType
  private overrides = new Map<string, string>(); // partName → contentType

  /** Add a Default entry (by file extension, without leading dot). */
  addDefault(extension: string, contentType: string): this {
    this.defaults.set(extension.toLowerCase(), contentType);
    return this;
  }

  /** Add an Override entry (by absolute part URI, e.g. "/word/document.xml"). */
  addOverride(partName: string, contentType: string): this {
    this.overrides.set(partName, contentType);
    return this;
  }

  /** Serialize to [Content_Types].xml string. */
  serialize(): string {
    const children: XmlElement[] = [];

    for (const [ext, ct] of this.defaults) {
      children.push(el('Default', { Extension: ext, ContentType: ct }));
    }
    for (const [part, ct] of this.overrides) {
      children.push(el('Override', { PartName: part, ContentType: ct }));
    }

    const root = el(
      'Types',
      { xmlns: NS.CONTENT_TYPES },
      ...children,
    );

    return xmlDocument(root);
  }
}
