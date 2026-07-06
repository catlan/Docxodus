import { describe, expect, test } from 'vitest';
import { IrTableDiffer, type IrDocument, type IrTable } from '../src/index.js';
import { fromBodyXml } from './helpers/ir-test-documents.js';

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const tableXml = (...rows: string[]) => `<w:tbl><w:tblPr/><w:tblGrid/>${rows.join('')}</w:tbl>`;
const fromXml = (body: string): IrDocument => fromBodyXml(body);

function onlyTable(doc: IrDocument): IrTable {
  const table = doc.body.blocks.find((b) => b.kind === 'table');
  if (!table || table.kind !== 'table') throw new Error('expected table');
  return table;
}

describe('IrTableDiffer', () => {
  test('cell text edit surfaces as token diff in that cell', () => {
    const left = onlyTable(fromXml(tableXml(row(cell('alpha one'), cell('beta two')), row(cell('gamma three'), cell('delta four')))));
    const right = onlyTable(fromXml(tableXml(row(cell('alpha one'), cell('beta two')), row(cell('gamma three'), cell('delta EDITED')))));
    const table = IrTableDiffer.diff(left, right);

    expect(table.rowOps[0]!.kind).toBe('EqualRow');
    const modRow = table.rowOps[1]!;
    expect(modRow.kind).toBe('ModifyRow');
    expect(modRow.cellOps).not.toBeNull();
    expect(modRow.cellOps![0]!.blockOps).toBeNull();
    const blockOp = modRow.cellOps![1]!.blockOps![0]!;
    expect(blockOp.kind).toBe('ModifyBlock');
    expect(blockOp.tokenDiff).not.toBeNull();
    const tokenDiff = blockOp.tokenDiff as { ops: ReadonlyArray<{ kind: string }> };
    expect(tokenDiff.ops.some((o) => o.kind === 'Insert' || o.kind === 'Delete')).toBe(true);
    expect(tokenDiff.ops.some((o) => o.kind === 'Equal')).toBe(true);
  });

  test('row inserted and deleted becomes one modified row between anchors', () => {
    const left = onlyTable(fromXml(tableXml(row(cell('keep me')), row(cell('delete me')), row(cell('also keep')))));
    const right = onlyTable(fromXml(tableXml(row(cell('keep me')), row(cell('brand new')), row(cell('also keep')))));
    const rowOps = IrTableDiffer.diff(left, right).rowOps;
    expect(rowOps.filter((o) => o.kind === 'EqualRow')).toHaveLength(2);
    expect(rowOps.filter((o) => o.kind === 'ModifyRow')).toHaveLength(1);
  });

  test('row only added', () => {
    const left = onlyTable(fromXml(tableXml(row(cell('one')), row(cell('two')))));
    const right = onlyTable(fromXml(tableXml(row(cell('one')), row(cell('two')), row(cell('three')))));
    const rowOps = IrTableDiffer.diff(left, right).rowOps;
    expect(rowOps.filter((o) => o.kind === 'EqualRow')).toHaveLength(2);
    expect(rowOps.filter((o) => o.kind === 'InsertRow')).toHaveLength(1);
    expect(rowOps.filter((o) => o.kind === 'DeleteRow')).toHaveLength(0);
  });

  test('deterministic table diff', () => {
    const left = onlyTable(fromXml(tableXml(row(cell('a'), cell('b')), row(cell('c'), cell('d')))));
    const right = onlyTable(fromXml(tableXml(row(cell('a'), cell('b')), row(cell('c'), cell('D-edited')))));
    expect(IrTableDiffer.diff(left, right)).toEqual(IrTableDiffer.diff(left, right));
  });
});
