// Port of Docxodus/Ir/IrAnchor.cs — the stable identity of an
// addressable IR node, rendered `kind:scope:unid` (the markdown
// projection's public anchor grammar).

export type IrAnchorKind =
  | 'p'
  | 'h'
  | 'li'
  | 'tbl'
  | 'tr'
  | 'tc'
  | 'cmt'
  | 'fn'
  | 'en'
  | 'img'
  | 'drw'
  | 'sec'
  | 'unk';

const KINDS: ReadonlySet<string> = new Set([
  'p', 'h', 'li', 'tbl', 'tr', 'tc', 'cmt', 'fn', 'en', 'img', 'drw', 'sec', 'unk',
]);

export interface IrAnchor {
  readonly kind: IrAnchorKind;
  readonly scope: string;
  readonly unid: string;
}

export const irAnchor = (
  kind: IrAnchorKind,
  scope: string,
  unid: string,
): IrAnchor => ({ kind, scope, unid });

export const anchorToString = (anchor: IrAnchor): string =>
  `${anchor.kind}:${anchor.scope}:${anchor.unid}`;

/** Inverse of anchorToString. Throws on malformed input (C# parity). */
export const anchorFromString = (value: string): IrAnchor => {
  const [kind, scope, unid, ...rest] = value.split(':');
  if (!kind || !scope || !unid || rest.length > 0 || !KINDS.has(kind)) {
    throw new Error(`Malformed IR anchor: '${value}'.`);
  }
  return { kind: kind as IrAnchorKind, scope, unid };
};
