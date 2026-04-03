/**
 * Hyperlinks — external URLs and internal bookmark anchors.
 *
 * Source: ECMA-376 5th ed. Part1/wml.xsd
 *   CT_Hyperlink  (line 1218) — hyperlink element in paragraph content
 *   CT_Rel        (line 1196) — base type with r:id attribute
 *
 * Traceability: SPEC_TRACEABILITY.md — WmlHyperlink, createHyperlink
 *
 * Key spec constraints:
 * - External hyperlink: r:id attribute points to a relationship with type "hyperlink"
 *   and TargetMode="External"
 * - Internal anchor: w:anchor attribute points to a bookmark in the document
 * - Hyperlink is paragraph-level content (EG_PContent), not run content
 * - Contains runs (w:r) as children
 *
 * Design: The package exposes generic hyperlink primitives.
 * Applications decide link destinations and display text.
 * No URL or bookmark is hardcoded.
 */

import type { WmlRun, WmlRunProperties } from './types';

// ---------------------------------------------------------------------------
// Type definitions
// ---------------------------------------------------------------------------

/**
 * Source: CT_Hyperlink (wml.xsd line 1218)
 * A hyperlink — either external (URL via r:id) or internal (bookmark anchor).
 * This is paragraph-level content, not run content.
 */
export interface WmlHyperlink {
  _type: 'hyperlink';
  /** Relationship id for external hyperlinks (set via builder.addHyperlinkRelationship()). */
  rId?: string;
  /** Bookmark anchor for internal hyperlinks. Mutually exclusive with rId. */
  anchor?: string;
  /** Run content displayed as the link text. */
  runs: WmlRun[];
  /** Optional tooltip shown on hover. */
  tooltip?: string;
  /** If true, the hyperlink history is not updated. */
  history?: boolean;
}

/**
 * Source: CT_Hyperlink + convenience wrapper
 * Parameters for creating a hyperlink.
 */
export interface WmlHyperlinkParams {
  /** External URL (e.g. "https://example.com"). Mutually exclusive with anchor. */
  url?: string;
  /** Internal bookmark name. Mutually exclusive with url. */
  anchor?: string;
  /** Display text for the link. */
  text: string;
  /** Optional run formatting applied to the link text. */
  rPr?: WmlRunProperties;
  /** Optional hover tooltip. */
  tooltip?: string;
}

// ---------------------------------------------------------------------------
// Factory helpers
// ---------------------------------------------------------------------------

/**
 * Create a hyperlink paragraph with a single text run.
 * For multi-run content, construct WmlHyperlink directly and use paragraph([hyperlinkRun]).
 *
 * @param params - Hyperlink parameters (url or anchor, text, optional formatting)
 * @returns A WmlHyperlink element — pass to paragraph() or add as block content
 *
 * @example
 * // External link
 * const link = createHyperlink({ url: 'https://example.com', text: 'Click here' });
 * builder.addBlock(paragraph([link]));
 *
 * // Internal bookmark link
 * const link = createHyperlink({ anchor: 'MyBookmark', text: 'Go to section' });
 * builder.addBlock(paragraph([link]));
 */
export function createHyperlink(params: WmlHyperlinkParams): WmlHyperlink {
  const rPr: WmlRunProperties = {
    ...params.rPr,
    // Default hyperlink styling if not overridden
    u: params.rPr?.u ?? 'single',
    color: params.rPr?.color ?? { val: '0563C1' },
  };

  return {
    _type: 'hyperlink',
    runs: [{ _type: 'run', rPr, content: [{ _type: 'text', text: params.text, space: 'preserve' }] }],
    tooltip: params.tooltip,
  };
}
