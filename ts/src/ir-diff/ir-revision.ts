export type IrRevisionType = 'Inserted' | 'Deleted' | 'Moved' | 'FormatChanged';

export type IrFormatChangeScope = 'Run' | 'Paragraph' | 'TableCell' | 'TableRow' | 'Table' | 'Section';

export interface IrFormatChangeDetails {
  readonly oldProperties: Readonly<Record<string, string>>;
  readonly newProperties: Readonly<Record<string, string>>;
  readonly changedPropertyNames: ReadonlyArray<string>;
  readonly scope: IrFormatChangeScope;
}

export interface IrRevision {
  readonly type: IrRevisionType;
  readonly text: string;
  readonly author: string;
  readonly date: string;
  readonly moveGroupId?: number | null;
  readonly isMoveSource?: boolean | null;
  readonly formatChange?: IrFormatChangeDetails | null;
  readonly leftAnchor?: string | null;
  readonly rightAnchor?: string | null;
}

export interface IrRevisionWire {
  readonly revisionType: IrRevisionType;
  readonly text: string;
  readonly author: string;
  readonly date: string;
  readonly moveGroupId: number | null;
  readonly isMoveSource: boolean | null;
  readonly formatChange: IrFormatChangeWire | null;
  readonly leftAnchor: string | null;
  readonly rightAnchor: string | null;
}

export interface IrFormatChangeWire {
  readonly oldProperties: Readonly<Record<string, string>>;
  readonly newProperties: Readonly<Record<string, string>>;
  readonly changedPropertyNames: ReadonlyArray<string>;
  readonly scope: 'run' | 'paragraph' | 'tableCell' | 'tableRow' | 'table' | 'section';
}

export function formatChangeScopeWire(scope: IrFormatChangeScope): IrFormatChangeWire['scope'] {
  switch (scope) {
    case 'Paragraph': return 'paragraph';
    case 'TableCell': return 'tableCell';
    case 'TableRow': return 'tableRow';
    case 'Table': return 'table';
    case 'Section': return 'section';
    case 'Run': return 'run';
  }
}

export function revisionToWire(revision: IrRevision): IrRevisionWire {
  return {
    revisionType: revision.type,
    text: revision.text,
    author: revision.author,
    date: revision.date,
    moveGroupId: revision.moveGroupId ?? null,
    isMoveSource: revision.isMoveSource ?? null,
    formatChange: revision.formatChange
      ? {
          oldProperties: revision.formatChange.oldProperties,
          newProperties: revision.formatChange.newProperties,
          changedPropertyNames: revision.formatChange.changedPropertyNames,
          scope: formatChangeScopeWire(revision.formatChange.scope),
        }
      : null,
    leftAnchor: revision.leftAnchor ?? null,
    rightAnchor: revision.rightAnchor ?? null,
  };
}

export function writeIrRevisionsJson(revisions: ReadonlyArray<IrRevision>): string {
  return JSON.stringify({ revisions: revisions.map(revisionToWire) });
}
