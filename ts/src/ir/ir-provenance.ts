// Port of Docxodus/Ir/IrProvenance.cs

import type { XElement } from '../xml/xelement.js';

/** Back-reference from an IR node to source OOXML; equality-neutral by convention in TS. */
export interface IrProvenance {
  /** The source OOXML element this IR node was read from, if known. */
  readonly element: XElement | null;
  /** The URI of the source part, if known. C# `Uri` ports to string. */
  readonly partUri: string | null;
  /** True when a block was delivered by an unwrapped block-level `w:sdt`; equality-neutral. */
  readonly fromBlockSdt: boolean;
}

/** Shared provenance carrying no element or part. */
export const emptyIrProvenance: IrProvenance = {
  element: null,
  partUri: null,
  fromBlockSdt: false,
};
