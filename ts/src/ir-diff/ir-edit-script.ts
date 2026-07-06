// Port of Docxodus/Ir/Diff/IrEditScript.cs data surface.

export type IrEditOpKind =
  | 'EqualBlock'
  | 'FormatOnlyBlock'
  | 'ModifyBlock'
  | 'InsertBlock'
  | 'DeleteBlock'
  | 'MoveBlock'
  | 'MoveModifyBlock'
  | 'SplitBlock'
  | 'MergeBlock';

export interface IrEditOp {
  readonly kind: IrEditOpKind;
  readonly leftAnchor: string | null;
  readonly rightAnchor: string | null;
  readonly tokenDiff: unknown | null;
  readonly moveGroupId: number | null;
  readonly isMoveSource: boolean | null;
  readonly tableDiff?: unknown | null;
  readonly textboxDiffs?: ReadonlyArray<IrTextboxDiff> | null;
  readonly splitMergeAnchors?: ReadonlyArray<string> | null;
  readonly segmentDiffs?: ReadonlyArray<unknown> | null;
}

export interface IrTextboxDiff {
  readonly ops: ReadonlyArray<IrEditOp>;
}

export interface IrNoteDiff {
  readonly kind: 'Footnote' | 'Endnote';
  readonly noteId: string;
  readonly ops: ReadonlyArray<IrEditOp>;
  readonly leftNoteId: string | null;
}

export interface IrHeaderFooterDiff {
  readonly isHeader: boolean;
  readonly kind: string;
  readonly sectionIndex: number;
  readonly scopeName: string;
  readonly leftScopeName: string | null;
  readonly leftPartUri: string | null;
  readonly rightPartUri: string | null;
  readonly ops: ReadonlyArray<IrEditOp>;
}

export type IrRowOpKind = 'EqualRow' | 'ModifyRow' | 'InsertRow' | 'DeleteRow' | 'MovedRow';

export interface IrRowOp {
  readonly kind: IrRowOpKind;
  readonly leftRowAnchor: string | null;
  readonly rightRowAnchor: string | null;
  readonly cellOps: ReadonlyArray<IrCellOp> | null;
  readonly moveGroupId: number | null;
  readonly isMoveSource: boolean | null;
}

export interface IrCellOp {
  readonly leftCellAnchor: string | null;
  readonly rightCellAnchor: string | null;
  readonly blockOps: ReadonlyArray<IrEditOp> | null;
}

export interface IrTableDiff {
  readonly rowOps: ReadonlyArray<IrRowOp>;
}

export interface IrEditScript {
  readonly ops: ReadonlyArray<IrEditOp>;
  readonly noteOps?: ReadonlyArray<IrNoteDiff>;
  readonly headerFooterOps?: ReadonlyArray<IrHeaderFooterDiff>;
}
