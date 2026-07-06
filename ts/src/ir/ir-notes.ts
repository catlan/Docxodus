import type { IrAnchor } from './ir-anchor.js';
import type { IrBlock } from './ir-blocks.js';
import type { IrScope } from './ir-reader.js';

export interface IrNoteStore {
  readonly notes: ReadonlyMap<string, IrScope>;
  readonly noteUnids: ReadonlyMap<string, string>;
}

export interface IrHeaderFooterRef {
  readonly sectionIndex: number;
  readonly kind: IrHeaderFooterKind;
}

export type IrHeaderFooterKind = 'Default' | 'First' | 'Even';

export interface IrHeaderFooter {
  readonly scopeName: string;
  readonly kind: IrHeaderFooterKind;
  readonly scope: IrScope;
  readonly references: ReadonlyArray<IrHeaderFooterRef>;
}

export interface IrCommentStore {
  readonly comments: ReadonlyArray<IrComment>;
  readonly partUri: string | null;
}

export interface IrComment {
  readonly anchor: IrAnchor;
  readonly author: string;
  readonly initials: string | null;
  readonly date: string | null;
  readonly blocks: ReadonlyArray<IrBlock>;
  readonly targets: ReadonlyArray<IrCommentTarget>;
}

export interface IrCommentTarget {
  readonly blockAnchor: IrAnchor;
  readonly startChar: number;
  readonly endChar: number;
}

export const EMPTY_NOTE_STORE: IrNoteStore = {
  notes: new Map(),
  noteUnids: new Map(),
};

export const EMPTY_COMMENT_STORE: IrCommentStore = {
  comments: [],
  partUri: null,
};
