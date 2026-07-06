// Port of Docxodus/Ir/IrFormats.cs

import type { IrHash } from './ir-hash.js';

/** Underline style (`w:u/@w:val`); `Other` covers unmodeled values. */
export type IrUnderlineKind =
  | 'Single'
  | 'Double'
  | 'Thick'
  | 'Dotted'
  | 'Dashed'
  | 'Wave'
  | 'Words'
  | 'None'
  | 'Other';

/** An underline: kind and optional color (`w:u/@w:color`, as written). */
export interface IrUnderline {
  readonly kind: IrUnderlineKind;
  readonly colorHex: string | null;
}

/** Vertical run alignment (`w:vertAlign`). */
export type IrVertAlign = 'Subscript' | 'Superscript';

/** Paragraph justification (`w:jc`); `Other` covers unmodeled values. */
export type IrJustification =
  | 'Left'
  | 'Center'
  | 'Right'
  | 'Both'
  | 'Distribute'
  | 'Other';

/** Line-spacing rule (`w:spacing/@w:lineRule`). */
export type IrLineSpacingRule = 'Auto' | 'AtLeast' | 'Exact';

/** Line spacing: value in twips plus interpreting rule. */
export interface IrLineSpacing {
  readonly valueTwips: number;
  readonly rule: IrLineSpacingRule;
}

/** Break kind (`w:br/@w:type`). */
export type IrBreakKind = 'Line' | 'Page' | 'Column';

/** Note kind: footnote or endnote. */
export type IrNoteKind = 'Footnote' | 'Endnote';

/** Cell vertical-merge state (`w:vMerge`). */
export type IrVMerge = 'None' | 'Restart' | 'Continue';

/** Direct section properties (`w:sectPr`). Unmodeled children fold into `unmodeledDigest`. */
export interface IrSectionFormat {
  readonly pageWidthTwips: number | null;
  readonly pageHeightTwips: number | null;
  readonly landscape: boolean | null;
  readonly marginTopTwips: number | null;
  readonly marginBottomTwips: number | null;
  readonly marginLeftTwips: number | null;
  readonly marginRightTwips: number | null;
  /** `w:type/@w:val` as written. */
  readonly sectionType: string | null;
  /** Digest of unmodeled `w:sectPr` children so changes still flip the format fingerprint. */
  readonly unmodeledDigest: IrHash;
}

/** Direct run properties (`w:rPr`), stored exactly as written. */
export interface IrRunFormat {
  readonly styleId: string | null;
  readonly bold: boolean | null;
  readonly italic: boolean | null;
  readonly underline: IrUnderline | null;
  readonly strike: boolean | null;
  readonly doubleStrike: boolean | null;
  readonly vertAlign: IrVertAlign | null;
  /** `w:rFonts/@w:ascii` as written; theme-resolved only in effective formats. */
  readonly fontAscii: string | null;
  readonly sizeHalfPoints: number | null;
  /** `w:color/@w:val` as written, including literal "auto". */
  readonly colorHex: string | null;
  readonly highlight: string | null;
  readonly caps: boolean | null;
  readonly smallCaps: boolean | null;
  readonly vanish: boolean | null;
  /** Digest of unmodeled `w:rPr` children so changes still flip the format fingerprint. */
  readonly unmodeledDigest: IrHash;
}

/** Direct paragraph properties (`w:pPr`), stored exactly as written. */
export interface IrParaFormat {
  readonly styleId: string | null;
  readonly justification: IrJustification | null;
  readonly indentLeftTwips: number | null;
  readonly indentRightTwips: number | null;
  /** First-line indent in twips; negative means hanging indent. */
  readonly indentFirstLineTwips: number | null;
  readonly spacingBeforeTwips: number | null;
  readonly spacingAfterTwips: number | null;
  readonly lineSpacing: IrLineSpacing | null;
  readonly outlineLevel: number | null;
  readonly keepNext: boolean | null;
  readonly keepLines: boolean | null;
  readonly pageBreakBefore: boolean | null;
  /** Direct list membership (`w:numPr/w:numId/@w:val`) as written. */
  readonly numId: number | null;
  /** Direct list level (`w:numPr/w:ilvl/@w:val`) as written. */
  readonly ilvl: number | null;
  /** Digest of unmodeled `w:pPr` children so changes still flip the format fingerprint. */
  readonly unmodeledDigest: IrHash;
}

/** List membership facts for a paragraph; null abstractNumId means "not yet resolved". */
export interface IrListInfo {
  readonly numId: number;
  readonly abstractNumId: number | null;
  readonly ilvl: number;
  readonly numberFormat: string;
  readonly startOverride: number | null;
  readonly fromStyle: boolean;
}
