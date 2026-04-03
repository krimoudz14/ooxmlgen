/**
 * OPC Relationship manager.
 *
 * Traceability: SPEC_TRACEABILITY.md — RelationshipManager, OpcRelationship
 * Source: 5TH EDITION PART2/opc-relationships.xsd
 *   CT_Relationships: root element of .rels files
 *   CT_Relationship: Id (xsd:ID, required), Type (URI, required),
 *                    Target (URI, required), TargetMode (optional)
 *
 * Constraint: Id MUST be unique within a .rels file (xsd:ID).
 * Convention: rId1, rId2, ... (matches Word's own output)
 */

import { el, xmlDocument } from '../shared/xml';
import { NS } from '../shared/namespaces';

/** Source: CT_Relationship / ST_TargetMode */
export type OpcTargetMode = 'Internal' | 'External';

/** Source: CT_Relationship */
export interface OpcRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: OpcTargetMode;
}

export class RelationshipManager {
  private rels: OpcRelationship[] = [];
  private counter = 0;

  /** Add a relationship and return its generated Id. */
  add(type: string, target: string, targetMode?: OpcTargetMode): string {
    this.counter++;
    const id = `rId${this.counter}`;
    this.rels.push({ id, type, target, targetMode });
    return id;
  }

  /** Add a relationship with an explicit Id. */
  addWithId(id: string, type: string, target: string, targetMode?: OpcTargetMode): void {
    this.rels.push({ id, type, target, targetMode });
  }

  getAll(): readonly OpcRelationship[] {
    return this.rels;
  }

  /** Serialize to .rels XML string. */
  serialize(): string {
    const children = this.rels.map(r =>
      el('Relationship', {
        Id: r.id,
        Type: r.type,
        Target: r.target,
        ...(r.targetMode ? { TargetMode: r.targetMode } : {}),
      }),
    );

    const root = el(
      'Relationships',
      { xmlns: NS.RELATIONSHIPS },
      ...children,
    );

    return xmlDocument(root);
  }
}
