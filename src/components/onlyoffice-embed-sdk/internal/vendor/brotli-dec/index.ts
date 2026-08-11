/** Vendored Brotli.js (MIT) with Google-derived Apache-2.0 code. See the package THIRD_PARTY_NOTICES. */
import { BrotliDecompressBuffer } from "./dec/decode";

export function brotliDecompress(input: Uint8Array): Uint8Array {
  return BrotliDecompressBuffer(input) as Uint8Array;
}
