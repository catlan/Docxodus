# IR-Powered DOCX Editor — Roadmap

Companion to `ir_editor_feasibility.md` (which records the verdict, architecture, and
PoC results). This is the **sequenced, prioritized** plan for turning the proven
foundation + MVP into a complete editor. Supersedes the scattered "Still Plan 2" notes.

Status (branch `feat/ir-editor-feasibility-poc`, PR #234): **foundation + MVP shipped and
proven; M1 (rich in-block editing) and M2 (structural editing) done; runnable demo at
`npm/examples/editor.html` (`npm run demo`).** M3 (worker offload) is next.

## Architecture invariants (do not break)

1. **Model-of-record = the live OOXML in `DocxSession`** (lossless `Save`). The IR is
   read-only and has no IR→OOXML writer; never make the IR or the DOM the source of truth.
2. **Addressing = the shared `{#kind:scope:unid}` anchor system.** `convertDocxToHtml`
   (stampAnchors) ↔ `DocxSession` ↔ `RenderBlock` all use one Unid scheme; keep it that way.
3. **Render is a projection; patch incrementally.** An edit goes through `DocxSession` by
   anchor, then only the changed block re-renders (`RenderBlockHtml`). Never round-trip the
   whole doc through `convertDocxToHtml` per edit.
4. **Untouched content stays byte-faithful on save.** Edits may simplify the *edited* block
   (within the markdown subset), but must never degrade blocks the user didn't touch.

## Shipped (foundation + MVP)

- C#: `WmlToHtmlConverterSettings.StampAnchors`; `HtmlConversionOps.RenderBlockHtml`
  (stateless + session-attached); `DocxSession.LiveDocument`.
- WASM/npm: `RenderBlockHtml`, `stampAnchors`, `renderBlockHtml()`, `DocxSession.renderBlock()`.
- `DocxEditor` (pure TS): faithful render → editable paragraphs/headings → commit via
  `DocxSession` → incremental re-render → lossless save; `{ paginated: true }` page boxes.
- Tests: C# `HCO050`/`HCO052`; browser `render-block.spec.ts`, `editor.spec.ts`.

## Milestones (priority order = impact)

### M1 — Rich in-block editing (preserve inline formatting)  · effort M–L · ✅ **DONE**
**Problem:** `commitBlock` replaced an edited block from `el.textContent` (plain text), so
editing a formatted paragraph destroyed its bold/italic/links — the biggest correctness trap.
**Shipped:** `serializeInlineMarkdown(block)` (exported from `npm/src/editor.ts`) walks the
edited block's DOM and emits the projector's markdown subset, detecting emphasis via
`getComputedStyle` (font-weight/font-style) and links via `href`, merging adjacent
same-format runs; `commitBlock` sends that markdown to `ReplaceText` instead of plain text.
Test `editor.spec.ts` "M1: editing preserves inline formatting" edits a block with
bold/italic/link and confirms `**…**` / `*…*` / `[…](…)` survive save+reopen. Formatting the
markdown subset cannot express (size/color) is still dropped on an *edited* block; a future
pass can use finer-grained `ReplaceTextAtSpan`/`ApplyFormat`. **Applying** new formatting
(toolbar) is M5.

### M2 — Structural editing via keyboard  · effort M · ✅ **DONE**
**Problem:** no way to add/split/merge blocks from the UI; ops existed in `DocxSession` but
weren't wired.
**Shipped:** Enter at the caret → `SplitParagraph(anchor, offset)` (offset from a Selection
range); Backspace at block start → `MergeParagraphs(prev, this)`. A `keydown` handler on each
block intercepts both, flushes any uncommitted typing first (`syncBlock`), applies the op,
and reconciles the DOM from `EditResult.modified/created/removed` (re-render the affected
block(s), insert/remove nodes, update the `unid → fullId` map, place the caret). Test
`editor.spec.ts` "M2: split and merge" splits a block (+1), merges the halves back (−1, text
restored exactly), and round-trips through save. Insert-at-doc-start and block delete/reorder
remain follow-ups (Enter-split + Backspace-merge cover the core authoring loop).

### M3 — Worker offload  · effort M–L
**Problem:** the initial full convert (~0.7–2.4 s) and session ops run on the main thread →
the UI freezes on open and on big docs. (Per-edit is already ~10 ms.)
**Approach:** extend the Web Worker surface (`docxodus.worker.ts` / `worker-proxy.ts`) to
carry session open/edit/render-block/save, transfer bytes zero-copy; the main thread holds
only the DOM. Keep the synchronous `DocxEditor` API working by awaiting worker round-trips.
**Acceptance:** opening and editing a large doc never blocks the main thread > ~16 ms;
existing editor tests pass through the worker path.

### M4 — Re-paginate on edit  · effort M
**Problem:** in paginated mode an edited block can overflow its page box (the MVP patches in
place without reflowing).
**Approach:** after a commit in paginated mode, re-run pagination from the affected page
forward (staging originals are retained, so a scoped reflow is feasible); debounce.
**Acceptance:** an edit that grows a block past a page boundary reflows to a new page.

### M5 — Formatting toolbar + undo/redo  · effort S–M
**Approach:** bold/italic/style/list controls → `ApplyFormat`/`SetParagraphStyle`/
`SetListLevel`; Ctrl+Z/Y → `DocxSession.Undo/Redo` (+ re-render affected blocks). Mostly UI
glue over existing ops.
**Acceptance:** toolbar applies formatting to a selection; undo/redo round-trips an edit.

### M6 — Tracked-changes / review mode  · effort M
**Approach:** open the session with `TrackedChanges = RenderInline`; render `ins`/`del` with
author colors; serve the redline/review use case.
**Acceptance:** edits land as `w:ins`/`w:del` with author attribution, visible in the editor.

### M7 — Table-cell & table-structure editing  · effort M
**Approach:** `ReplaceCellContent` for cell text; row/col insert-delete and cell-merge via
`session.Raw.*` until first-class ops exist.
**Acceptance:** a cell's content edits and round-trips; tables are no longer read-only.

### M8 — React wrapper  · effort S
**Approach:** `useDocxEditor` hook + `<DocxEditor>` component over the pure-TS core, in
`npm/src/react.ts`.
**Acceptance:** a React app mounts the editor with one component.

### M9 — Single-block render fidelity  · effort M
**Approach:** copy image parts into the throwaway render doc; resolve list-numbering
continuation for a block rendered in isolation.
**Acceptance:** re-rendering an image-bearing or list-item block matches the full render.

## Recommended sequencing

**M1 + M2 done** — "make editing real" is complete (edits preserve formatting; Enter/Backspace
split/merge). A runnable demo exists (`npm run demo` → `editor.html`). **M3 next** (worker
offload) for responsiveness on large docs. M4–M9 sequence by target use case: authoring favors
M4/M5; review favors M6; broad fidelity favors M7/M9. M8 (React) any time.
