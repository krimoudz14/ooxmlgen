/**
 * OPC Package — wraps JSZip to produce a valid OOXML ZIP container.
 *
 * Traceability: SPEC_TRACEABILITY.md — OpcPackage
 * Source: 5TH EDITION PART2/opc-contentTypes.xsd, opc-relationships.xsd
 *
 * OPC rules implemented:
 * - [Content_Types].xml at package root
 * - /_rels/.rels at package root (package-level relationships)
 * - Part-level .rels at <dir>/_rels/<filename>.rels
 * - Every part must have a content type
 */

import JSZip from 'jszip';
import { ContentTypeManager } from './content-types';
import { RelationshipManager, OpcTargetMode } from './relationships';
import { CONTENT_TYPES } from '../shared/namespaces';

export class OpcPackage {
  private zip = new JSZip();
  private contentTypes = new ContentTypeManager();
  private packageRels = new RelationshipManager();
  private partRels = new Map<string, RelationshipManager>();

  constructor() {
    // Default content types required by OPC spec
    this.contentTypes.addDefault('rels', CONTENT_TYPES.RELS);
    this.contentTypes.addDefault('xml', CONTENT_TYPES.XML);
  }

  /** Add a part (file) to the package. */
  addPart(partPath: string, contentType: string, content: string | Uint8Array): void {
    // Strip leading slash for JSZip
    const zipPath = partPath.startsWith('/') ? partPath.slice(1) : partPath;
    this.zip.file(zipPath, content);
    this.contentTypes.addOverride(partPath, contentType);
  }

  /** Add a package-level relationship. Returns the generated rId. */
  addPackageRelationship(type: string, target: string, targetMode?: OpcTargetMode): string {
    return this.packageRels.add(type, target, targetMode);
  }

  /** Add a part-level relationship. Returns the generated rId. */
  addPartRelationship(partPath: string, type: string, target: string, targetMode?: OpcTargetMode): string {
    if (!this.partRels.has(partPath)) {
      this.partRels.set(partPath, new RelationshipManager());
    }
    return this.partRels.get(partPath)!.add(type, target, targetMode);
  }

  /** Build and return the DOCX as a Buffer. */
  async build(): Promise<Buffer> {
    // Write [Content_Types].xml
    this.zip.file('[Content_Types].xml', this.contentTypes.serialize());

    // Write package-level /_rels/.rels
    this.zip.file('_rels/.rels', this.packageRels.serialize());

    // Write part-level .rels files
    for (const [partPath, relMgr] of this.partRels) {
      const normalized = partPath.startsWith('/') ? partPath.slice(1) : partPath;
      const lastSlash = normalized.lastIndexOf('/');
      const dir = lastSlash >= 0 ? normalized.slice(0, lastSlash) : '';
      const filename = lastSlash >= 0 ? normalized.slice(lastSlash + 1) : normalized;
      const relsPath = dir ? `${dir}/_rels/${filename}.rels` : `_rels/${filename}.rels`;
      this.zip.file(relsPath, relMgr.serialize());
    }

    return this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
}
