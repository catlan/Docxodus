// Port of Docxodus/Ir/IrInlines.cs

import type { XName } from '../xml/xelement.js';
import type { IrAnchor } from './ir-anchor.js';
import type { IrBlock } from './ir-blocks.js';
import type { IrHash } from './ir-hash.js';
import type { IrBreakKind, IrNoteKind, IrRunFormat } from './ir-formats.js';

/** Base union for inline-level IR content (paragraph children). */
export type IrInline =
  | IrTextRun
  | IrTab
  | IrBreak
  | IrHyperlink
  | IrFieldRun
  | IrNoteRef
  | IrInlineImage
  | IrTextbox
  | IrOpaqueInline;

/** A run of literal text with direct run formatting. */
export interface IrTextRun {
  readonly kind: 'textRun';
  readonly text: string;
  readonly format: IrRunFormat;
  /** True when spliced from inline `w:sdt`/`w:smartTag`; equality-neutral in C#. */
  readonly fromInlineSdt: boolean;
}

/** A tab character (`w:tab`) carrying its containing run formatting. */
export interface IrTab {
  readonly kind: 'tab';
  readonly format: IrRunFormat;
}

/** A break (`w:br`) of the given kind. */
export interface IrBreak {
  readonly kind: 'break';
  readonly breakKind: IrBreakKind;
}

/** A hyperlink (`w:hyperlink`); exactly one of target/internalTarget is expected. */
export interface IrHyperlink {
  readonly kind: 'hyperlink';
  readonly target: string | null;
  readonly internalTarget: IrAnchor | null;
  readonly inlines: ReadonlyArray<IrInline>;
}

/** A field modeled as instruction plus cached result inlines. */
export interface IrFieldRun {
  readonly kind: 'fieldRun';
  readonly instruction: string;
  readonly cachedResult: ReadonlyArray<IrInline>;
  /** True for `w:fldSimple`; false for run-based field machinery. */
  readonly isSimpleField: boolean;
}

/** A footnote/endnote reference (`w:footnoteReference`/`w:endnoteReference`). */
export interface IrNoteRef {
  readonly kind: 'noteRef';
  readonly noteKind: IrNoteKind;
  readonly noteId: string;
}

/** Inline image: part URI, byte hash, EMU dimensions, alt text, and equality-neutral unid. */
export interface IrInlineImage {
  readonly kind: 'inlineImage';
  readonly partUri: string;
  readonly imageBytesHash: IrHash;
  readonly widthEmu: bigint;
  readonly heightEmu: bigint;
  readonly altText: string | null;
  /** Source `w:drawing` `pt:Unid`, equality-neutral in C#. */
  readonly unid: string | null;
}

/** A textbox body; inner blocks are fully modeled and compose by value. */
export interface IrTextbox {
  readonly kind: 'textbox';
  readonly blocks: ReadonlyArray<IrBlock>;
}

/** Unmodeled inline preserved opaquely by element name plus canonical XML hash. */
export interface IrOpaqueInline {
  readonly kind: 'opaqueInline';
  readonly elementName: XName;
  readonly canonicalHash: IrHash;
}
