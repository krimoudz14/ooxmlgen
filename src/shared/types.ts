/**
 * Shared OOXML primitive types.
 * Source: ECMA-376 5th ed. Part1/shared-commonSimpleTypes.xsd
 *         Part1/shared-relationshipReference.xsd
 */

/**
 * Measurement in twips (twentieths of a point).
 * 1 twip = 1/1440 inch = 1/20 pt.
 * Source: ST_TwipsMeasure
 */
export type Twips = number;

/**
 * 6-character hex RGB color string, e.g. "FF0000".
 * Source: ST_HexColorRGB
 */
export type HexColorRGB = string;

/**
 * Hex color: either "auto" or a 6-char RGB hex string.
 * Source: ST_HexColor
 */
export type HexColor = 'auto' | HexColorRGB;

/**
 * BCP47 language tag, e.g. "en-US".
 * Source: ST_Lang
 */
export type LanguageTag = string;

/**
 * GUID in the form {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}.
 * Source: ST_Guid
 */
export type Guid = string;

/**
 * Conformance class for the document.
 * Source: ST_ConformanceClass
 */
export type ConformanceClass = 'strict' | 'transitional';

/**
 * Relationship ID string (xsd:ID in .rels files).
 * Source: ST_RelationshipId
 */
export type RelationshipId = string;

/**
 * Font size in half-points (e.g. 24 = 12pt).
 * Source: ST_HpsMeasure
 */
export type HalfPoints = number;

/**
 * Signed twips measure (can be negative for offsets).
 * Source: ST_SignedTwipsMeasure
 */
export type SignedTwips = number;

/** Helper: convert points to twips */
export function ptToTwips(pt: number): Twips {
  return Math.round(pt * 20);
}

/** Helper: convert inches to twips */
export function inToTwips(inches: number): Twips {
  return Math.round(inches * 1440);
}

/** Helper: convert cm to twips */
export function cmToTwips(cm: number): Twips {
  return Math.round(cm * 567);
}

/** Helper: convert points to half-points */
export function ptToHalfPt(pt: number): HalfPoints {
  return Math.round(pt * 2);
}
