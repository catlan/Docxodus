// Port of Docxodus/Ir/Diff/IrBlockAlignment.cs

import type { IrBlock } from '../ir/ir-blocks.js';

export type IrAlignmentKind =
  | 'Unchanged'
  | 'FormatOnly'
  | 'Modified'
  | 'Moved'
  | 'MovedModified'
  | 'Inserted'
  | 'Deleted'
  | 'Split'
  | 'Merge';

export interface IrAlignedBlock {
  readonly kind: IrAlignmentKind;
  readonly left: IrBlock | null;
  readonly right: IrBlock | null;
  readonly multiBlocks: ReadonlyArray<IrBlock> | null;
}

export interface IrBlockAlignment {
  readonly entries: ReadonlyArray<IrAlignedBlock>;
}
