// docxodus-ts — TypeScript port of the DocxDiff IR engine.
// Port status: v2.1 foundation (XML substrate, hashing, anchors,
// deterministic unids). The IR reader, tokenizer, and diff pipeline
// land in subsequent phases; see diff.tools docs/word-diff-plan.md §v2.

export * from './xml/xelement.js';
export * from './xml/xwriter.js';
export * from './ir/names.js';
export * from './ir/short-hash.js';
export * from './ir/ir-hash.js';
export * from './ir/ir-anchor.js';
export * from './ir/unid-helper.js';
export * from './ir/ir-formats.js';
export * from './ir/ir-provenance.js';
export * from './ir/ir-inlines.js';
export * from './ir/ir-blocks.js';
export * from './ir/ir-notes.js';
export * from './ir-diff/ir-diff-token.js';
export * from './ir-diff/ir-diff-settings.js';
export * from './ir-diff/ir-diff-tokenizer.js';
export * from './ir-diff/ir-token-diff.js';
export * from './ir-diff/ir-token-differ.js';
export * from './ir-diff/ir-edit-script.js';
export * from './ir-diff/ir-edit-script-builder.js';
export * from './ir-diff/ir-edit-script-json.js';
export * from './ir-diff/ir-revision.js';
export * from './ir-diff/ir-revision-renderer.js';
export * from './ir-diff/ir-table-differ.js';
export * from './ir-diff/ir-split-segmenter.js';
export * from './ir-diff/ir-block-alignment.js';
export * from './ir-diff/ir-block-similarity.js';
export * from './ir-diff/ir-block-aligner.js';
export * from './ir-diff/ir-modeled-format.js';
export * from './ir/ir-content-hash-builder.js';
export * from './ir/ir-canonicalize.js';
export * from './ir/kind-for.js';
export * from './ir/ir-reader.js';
