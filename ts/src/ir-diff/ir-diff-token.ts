// Port of Docxodus/Ir/Diff/IrDiffToken.cs

import type { IrRunFormat } from '../ir/ir-formats.js';

/** Diff token kind; mirrors the §6.1 content-hash stream granularity. */
export type IrDiffTokenKind =
  | 'Word'
  | 'Separator'
  | 'Tab'
  | 'Break'
  | 'NoteRef'
  | 'Image'
  | 'Opaque'
  | 'Textbox';

/** One diff-time token over an IR paragraph. */
export interface IrDiffToken {
  readonly kind: IrDiffTokenKind;
  /** Raw source text, empty for non-text kinds. */
  readonly text: string;
  /** Normalized equality key. */
  readonly matchKey: string;
  /** Half-open offsets counting only emitted IrTextRun characters. */
  readonly startChar: number;
  readonly endChar: number;
  /** Governing run format, null for non-run kinds. */
  readonly format: IrRunFormat | null;
}
