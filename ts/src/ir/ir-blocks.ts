// Port of Docxodus/Ir/IrBlocks.cs

import type { XName } from '../xml/xelement.js';
import type { IrAnchor } from './ir-anchor.js';
import type {
  IrListInfo,
  IrParaFormat,
  IrSectionFormat,
  IrVMerge,
} from './ir-formats.js';
import type { IrHash } from './ir-hash.js';
import type { IrInline } from './ir-inlines.js';
import { emptyIrProvenance, type IrProvenance } from './ir-provenance.js';

/** Shared block fields: stable anchor, content hash, format fingerprint, and equality-neutral source. */
export interface IrBlockBase {
  readonly anchor: IrAnchor;
  readonly contentHash: IrHash;
  readonly formatFingerprint: IrHash;
  /** Back-reference to source OOXML; equality-neutral. */
  readonly source: IrProvenance;
}

export type IrBlock = IrParagraph | IrTable | IrSectionBreak | IrOpaqueBlock;

/** A paragraph: direct formatting, optional list membership, and inline children. */
export interface IrParagraph extends IrBlockBase {
  readonly kind: 'paragraph';
  /** Direct paragraph formatting; cascade resolution is a computed view. */
  readonly format: IrParaFormat;
  readonly list: IrListInfo | null;
  readonly inlines: ReadonlyArray<IrInline>;
  /**
   * Resolved auto-number marker Word would render. Diff code must key alignment/equality on
   * contentHash/formatFingerprint, never record equality.
   */
  readonly resolvedListMarker: string | null;
  /** Anchor of inline `w:pPr/w:sectPr` section transition, if present. */
  readonly inlineSectionBreakAnchor: IrAnchor | null;
  /** Modeled format of inline `w:pPr/w:sectPr`, folded into formatFingerprint. */
  readonly inlineSectionFormat: IrSectionFormat | null;
  /** Canonical paragraph-property projection for composite attribution; not in contentHash/fingerprint. */
  readonly pPrDigest: IrHash;
  /** Oracle structural list-item verdict used for markdown spacing. */
  readonly isListItemForLayout: boolean;
}

/** A table: rows plus table-property and grid digests. */
export interface IrTable extends IrBlockBase {
  readonly kind: 'table';
  readonly rows: ReadonlyArray<IrRow>;
  /** Canonical hash of the table's `w:tblPr` shell, folded into formatFingerprint. */
  readonly tblPrDigest: IrHash;
  /** Canonical hash of `w:tblGrid`, folded into formatFingerprint. */
  readonly tblGridDigest: IrHash;
}

/** A table row; source provenance is equality-neutral. */
export interface IrRow {
  readonly kind: 'row';
  readonly anchor: IrAnchor;
  readonly cells: ReadonlyArray<IrCell>;
  readonly contentHash: IrHash;
  readonly source: IrProvenance;
  /** Canonical row shell hash, folded into the table format fingerprint. */
  readonly trPrDigest: IrHash;
  /** Canonical `w:trPr` flattened-children digest; not in fingerprint. */
  readonly trPrShellDigest: IrHash;
  /** Canonical `w:tblPrEx` flattened-children digest; not in fingerprint. */
  readonly trPrExDigest: IrHash;
  /** True when delivered by table-level `w:sdt`; equality-participating. */
  readonly fromTableSdt: boolean;
}

/** A table cell: blocks plus grid span and vertical-merge state. */
export interface IrCell {
  readonly kind: 'cell';
  readonly anchor: IrAnchor;
  readonly blocks: ReadonlyArray<IrBlock>;
  readonly gridSpan: number;
  readonly vMerge: IrVMerge;
  readonly contentHash: IrHash;
  readonly source: IrProvenance;
  /** Canonical whole-`w:tcPr` shell hash, folded into contentHash. */
  readonly shellDigest: IrHash;
  /** Canonical flattened `w:tcPr` children digest; not in contentHash. */
  readonly tcPrShellDigest: IrHash;
  /** True when delivered by row-level `w:sdt`; equality-participating. */
  readonly fromRowSdt: boolean;
}

/** A section break carrying direct section formatting (`w:sectPr`). */
export interface IrSectionBreak extends IrBlockBase {
  readonly kind: 'sectionBreak';
  readonly format: IrSectionFormat;
}

/** Unmodeled block-level element preserved opaquely. */
export interface IrOpaqueBlock extends IrBlockBase {
  readonly kind: 'opaqueBlock';
  readonly elementName: XName;
}

export const withEmptyProvenance = <T extends Omit<IrBlockBase, 'source'>>(
  block: T,
): T & { readonly source: IrProvenance } => ({ ...block, source: emptyIrProvenance });
