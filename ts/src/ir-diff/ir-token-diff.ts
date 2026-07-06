// Port of Docxodus/Ir/Diff/IrTokenDiff.cs

export type IrTokenOpKind = 'Equal' | 'Insert' | 'Delete' | 'FormatChanged';

export interface IrTokenOp {
  readonly kind: IrTokenOpKind;
  readonly leftStart: number;
  readonly leftEnd: number;
  readonly rightStart: number;
  readonly rightEnd: number;
}

export interface IrTokenDiff {
  readonly ops: ReadonlyArray<IrTokenOp>;
}

export const tokenOp = (
  kind: IrTokenOpKind,
  leftStart: number,
  leftEnd: number,
  rightStart: number,
  rightEnd: number,
): IrTokenOp => ({ kind, leftStart, leftEnd, rightStart, rightEnd });

export const tokenOpLeftLength = (op: IrTokenOp): number => op.leftEnd - op.leftStart;
export const tokenOpRightLength = (op: IrTokenOp): number => op.rightEnd - op.rightStart;
