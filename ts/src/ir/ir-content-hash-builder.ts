// Port of Docxodus/Ir/IrHasher.cs → IrContentHashBuilder (spec §6.1):
// accumulates the canonical byte stream for a block's contentHash and
// SHA-256s it on build. Text is appended as UTF-8; non-text structure
// as sentinel/marker byte sequences that cannot collide with text
// (lead bytes 0x01/0x02 are forbidden in XML 1.0 content, so no
// XML-sourced string can contain them).

import { sha256 } from '@noble/hashes/sha2.js';
import { irHashFromBytes, irHashToBytes, type IrHash } from './ir-hash.js';

// Sentinel kinds (after a 0x01 lead byte) — non-text inlines.
export const SENTINEL_TAB = 0x01;
export const SENTINEL_LINE_BREAK = 0x02;
export const SENTINEL_PAGE_BREAK = 0x03;
export const SENTINEL_COLUMN_BREAK = 0x04;
export const SENTINEL_FOOTNOTE_REF = 0x05;
export const SENTINEL_ENDNOTE_REF = 0x06;
export const SENTINEL_IMAGE = 0x07;
// Hyperlink framing (§6.1 N14): target bracketed between these two, child
// inline bytes follow — linked text never collides with plain text.
export const SENTINEL_HYPERLINK = 0x08;
export const SENTINEL_HYPERLINK_TARGET_END = 0x09;
// Textbox framing (§6.1 M1.5): sentinel + each inner block's 32-byte
// contentHash in order.
export const SENTINEL_TEXTBOX = 0x0b;
export const SENTINEL_OPAQUE = 0x0f;

// Structure markers (after a 0x02 lead byte) — table structure.
export const STRUCTURE_ROW = 0x10;
export const STRUCTURE_CELL = 0x11;

const SENTINEL_LEAD = 0x01;
const STRUCTURE_LEAD = 0x02;

export class IrContentHashBuilder {
  private chunks: Uint8Array[] = [];
  private readonly encoder = new TextEncoder();

  appendText(text: string): void {
    if (text.length > 0) this.chunks.push(this.encoder.encode(text));
  }

  appendSentinel(kind: number): void {
    this.chunks.push(Uint8Array.of(SENTINEL_LEAD, kind));
  }

  appendStructure(marker: number): void {
    this.chunks.push(Uint8Array.of(STRUCTURE_LEAD, marker));
  }

  /** Append the 32 raw bytes of a hash (image / opaque / textbox block). */
  appendHash(hash: IrHash): void {
    this.chunks.push(irHashToBytes(hash));
  }

  build(): IrHash {
    let length = 0;
    for (const chunk of this.chunks) length += chunk.length;
    const buffer = new Uint8Array(length);
    let offset = 0;
    for (const chunk of this.chunks) {
      buffer.set(chunk, offset);
      offset += chunk.length;
    }
    return irHashFromBytes(sha256(buffer));
  }
}
