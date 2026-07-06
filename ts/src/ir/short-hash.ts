// Port of Docxodus/UnidHelper.cs ShortHash: the first `hexChars`
// lowercase hex characters of SHA-256(UTF-8(input)). @noble/hashes is
// synchronous and environment-independent (WebCrypto's digest is async,
// which would poison every call site; node:crypto is Node-only).

import { sha256 } from '@noble/hashes/sha2.js';

const HEX = '0123456789abcdef';

export function shortHash(input: string, hexChars: number): string {
  const digest = sha256(new TextEncoder().encode(input));
  const byteCount = hexChars / 2;
  let out = '';
  for (let i = 0; i < byteCount; i++) {
    const byte = digest[i]!;
    out += HEX[byte >> 4]! + HEX[byte & 0xf]!;
  }
  return out;
}
