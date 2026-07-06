// docxodus-ts — TypeScript port of the DocxDiff IR engine.
// Port status: v2.1 foundation (XML substrate, hashing, anchors,
// deterministic unids). The IR reader, tokenizer, and diff pipeline
// land in subsequent phases; see diff.tools docs/word-diff-plan.md §v2.

export * from './xml/xelement.js';
export * from './ir/names.js';
export * from './ir/short-hash.js';
export * from './ir/ir-hash.js';
export * from './ir/ir-anchor.js';
export * from './ir/unid-helper.js';
export * from './ir/ir-formats.js';
export * from './ir/ir-provenance.js';
export * from './ir/ir-inlines.js';
export * from './ir/ir-blocks.js';
export * from './ir-diff/ir-diff-token.js';
export * from './ir-diff/ir-diff-tokenizer.js';
