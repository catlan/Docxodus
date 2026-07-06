import type { IrEditOp, IrEditScript, IrHeaderFooterDiff, IrNoteDiff, IrTableDiff } from './ir-edit-script.js';
import type { IrTokenDiff, IrTokenOpKind } from './ir-token-diff.js';

export function writeIrEditScriptJson(script: IrEditScript): string {
  const root: Record<string, unknown> = {
    operations: script.ops.map(writeOpObject),
  };
  if (script.noteOps !== undefined && script.noteOps.length > 0) {
    root.noteOps = script.noteOps.map(writeNoteDiffObject);
  }
  if (script.headerFooterOps !== undefined && script.headerFooterOps.length > 0) {
    root.headerFooterOps = script.headerFooterOps.map(writeHeaderFooterDiffObject);
  }
  return JSON.stringify(root, null, 2);
}

function writeOpObject(op: IrEditOp): Record<string, unknown> {
  const obj: Record<string, unknown> = { kind: op.kind };
  if (op.leftAnchor !== null) obj.leftAnchor = op.leftAnchor;
  if (op.rightAnchor !== null) obj.rightAnchor = op.rightAnchor;
  if (op.moveGroupId !== null) obj.moveGroupId = op.moveGroupId;
  if (op.isMoveSource !== null) obj.isMoveSource = op.isMoveSource;
  if (op.tokenDiff !== null) obj.tokenDiff = writeTokenDiffObject(op.tokenDiff as IrTokenDiff);
  if (op.tableDiff !== undefined && op.tableDiff !== null) obj.tableDiff = writeTableDiffObject(op.tableDiff as IrTableDiff);
  if (op.textboxDiffs !== undefined && op.textboxDiffs !== null) {
    obj.textboxDiffs = op.textboxDiffs.map((d) => ({ ops: d.ops.map(writeOpObject) }));
  }
  if (op.splitMergeAnchors !== undefined && op.splitMergeAnchors !== null) obj.splitMergeAnchors = [...op.splitMergeAnchors];
  if (op.segmentDiffs !== undefined && op.segmentDiffs !== null) {
    obj.segmentDiffs = op.segmentDiffs.map((d) => writeTokenDiffObject(d as IrTokenDiff));
  }
  return obj;
}

function writeNoteDiffObject(diff: IrNoteDiff): Record<string, unknown> {
  const obj: Record<string, unknown> = {
    kind: diff.kind,
    noteId: diff.noteId,
  };
  if (diff.leftNoteId !== null) obj.leftNoteId = diff.leftNoteId;
  obj.ops = diff.ops.map(writeOpObject);
  return obj;
}

function writeHeaderFooterDiffObject(diff: IrHeaderFooterDiff): Record<string, unknown> {
  const obj: Record<string, unknown> = {
    isHeader: diff.isHeader,
    kind: diff.kind,
    sectionIndex: diff.sectionIndex,
    scope: diff.scopeName,
  };
  if (diff.leftScopeName !== null) obj.leftScope = diff.leftScopeName;
  if (diff.leftPartUri !== null) obj.leftPartUri = diff.leftPartUri;
  if (diff.rightPartUri !== null) obj.rightPartUri = diff.rightPartUri;
  obj.ops = diff.ops.map(writeOpObject);
  return obj;
}

function writeTableDiffObject(diff: IrTableDiff): Record<string, unknown> {
  return {
    rowOps: diff.rowOps.map((op) => {
      const obj: Record<string, unknown> = { kind: op.kind };
      if (op.leftRowAnchor !== null) obj.leftRowAnchor = op.leftRowAnchor;
      if (op.rightRowAnchor !== null) obj.rightRowAnchor = op.rightRowAnchor;
      if (op.moveGroupId !== null) obj.moveGroupId = op.moveGroupId;
      if (op.isMoveSource !== null) obj.isMoveSource = op.isMoveSource;
      if (op.cellOps !== null) {
        obj.cellOps = op.cellOps.map((cellOp) => {
          const cell: Record<string, unknown> = {};
          if (cellOp.leftCellAnchor !== null) cell.leftCellAnchor = cellOp.leftCellAnchor;
          if (cellOp.rightCellAnchor !== null) cell.rightCellAnchor = cellOp.rightCellAnchor;
          if (cellOp.blockOps !== null) cell.blockOps = cellOp.blockOps.map(writeOpObject);
          return cell;
        });
      }
      return obj;
    }),
  };
}

function writeTokenDiffObject(diff: IrTokenDiff): Record<string, unknown> {
  return {
    ops: diff.ops.map((op) => [
      tokenKindCode(op.kind),
      op.leftStart,
      op.leftEnd,
      op.rightStart,
      op.rightEnd,
    ]),
  };
}

function tokenKindCode(kind: IrTokenOpKind): number {
  switch (kind) {
    case 'Equal':
      return 0;
    case 'Insert':
      return 1;
    case 'Delete':
      return 2;
    case 'FormatChanged':
      return 3;
  }
}

export const IrEditScriptJson = {
  write: writeIrEditScriptJson,
};
