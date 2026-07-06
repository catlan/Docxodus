// Port of Docxodus/Ir/IrHash.cs — minimum TypeScript IR hash layer.

import { sha256 } from '@noble/hashes/sha2.js';

const HEX = '0123456789abcdef';
const IR_HASH_RE = /^[0-9a-f]{64}$/;

declare const irHashBrand: unique symbol;

/** A 32-byte SHA-256 digest rendered as 64 lowercase hex characters. */
export type IrHash = string & { readonly [irHashBrand]: true };

export function irHashFromHex(hex: string): IrHash {
  if (!IR_HASH_RE.test(hex)) {
    throw new Error(`Malformed IrHash hex: '${hex}'.`);
  }
  return hex as IrHash;
}

/** Compute the SHA-256 digest of UTF-8 text or raw bytes. */
export function irHashCompute(bytesOrString: string | Uint8Array): IrHash {
  const bytes =
    typeof bytesOrString === 'string'
      ? new TextEncoder().encode(bytesOrString)
      : bytesOrString;
  const digest = sha256(bytes);
  let out = '';
  for (const byte of digest) {
    out += HEX[byte >> 4]! + HEX[byte & 0xf]!;
  }
  return out as IrHash;
}

export const irHashEquals = (left: IrHash, right: IrHash): boolean => left === right;

/** The 32 raw digest bytes (big-endian hex order, matching toHex). */
export function irHashToBytes(hash: IrHash): Uint8Array {
  const bytes = new Uint8Array(32);
  for (let i = 0; i < 32; i++) {
    bytes[i] = parseInt(hash.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

/** Wrap a 32-byte digest as an IrHash (inverse of irHashToBytes). */
export function irHashFromBytes(bytes: Uint8Array): IrHash {
  if (bytes.length !== 32) {
    throw new Error(`IrHash requires 32 bytes, got ${bytes.length}.`);
  }
  let hex = '';
  for (const byte of bytes) hex += byte.toString(16).padStart(2, '0');
  return irHashFromHex(hex);
}
