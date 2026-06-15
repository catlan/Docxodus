/**
 * DocxEditor — a framework-agnostic, in-browser DOCX block editor.
 *
 * Architecture (see docs/architecture/ir_editor_feasibility.md, "Option B"):
 *   - model-of-record: a live DocxSession in WASM (lossless save);
 *   - rendering: WmlToHtmlConverter HTML (faithful) stamped with data-anchor;
 *   - editing: each block is contenteditable; on commit, the edit goes through
 *     DocxSession by anchor, then ONLY that block is re-rendered from the live
 *     session (session-attached RenderBlockHtml) and patched into the DOM.
 *
 * The IR/anchor system is the addressing spine; the live OOXML is the truth.
 * This is the pure-TypeScript core; a React wrapper can sit on top.
 *
 * MVP scope: per-block, commit-on-blur editing of paragraphs/headings. An edited
 * block's content is replaced from its plain text (inline formatting within an
 * edited block is not preserved — a documented MVP limit); UNTOUCHED blocks keep
 * full fidelity, and save() is lossless for them.
 */

import { paginateHtml } from "./pagination.js";

/** The subset of WASM bridge exports the editor needs (as exposed on `window.Docxodus`). */
export interface DocxEditorExports {
  DocxSessionBridge: {
    OpenSession: (bytes: Uint8Array, settingsJson: string) => number;
    CloseSession: (handle: number) => void;
    Project: (handle: number) => string;
    ReplaceText: (handle: number, anchor: string, md: string) => string;
    SplitParagraph: (handle: number, anchor: string, offset: number) => string;
    MergeParagraphs: (handle: number, first: string, second: string) => string;
    RenderBlockHtml: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    Save: (handle: number) => Uint8Array;
    Undo: (handle: number) => string;
    Redo: (handle: number) => string;
  };
  DocumentConverter: {
    ConvertDocxToHtmlComplete: (...args: any[]) => string;
  };
}

export interface DocxEditorOptions {
  /** CSS class prefix for rendered HTML. Default "docx-". */
  cssPrefix?: string;
  /** Fabricate CSS classes (vs inline styles). Default true. */
  fabricateClasses?: boolean;
  /** Make paragraph/heading blocks editable. Default true. */
  editable?: boolean;
  /** Render block-flow pages (page boxes via pagination.ts) vs a continuous view. Default false. */
  paginated?: boolean;
  /** Page render scale for paginated mode (1.0 = 100%). Default 1. */
  scale?: number;
  /** Called after a block edit commits (with the affected anchor). */
  onEdit?: (info: { anchorId: string; unid: string }) => void;
}

interface AnchorTargetLite {
  unid: string;
  kind: string;
  scope: string;
  textPreview?: string;
}

const EDITABLE_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6"]);

// ─── M1: inline HTML → markdown (preserve formatting on edit) ───────────────

interface InlineSeg {
  text: string;
  bold: boolean;
  italic: boolean;
  href: string | null;
}

function fontWeightIsBold(w: string): boolean {
  if (w === "bold" || w === "bolder") return true;
  const n = parseInt(w, 10);
  return !Number.isNaN(n) && n >= 600;
}

function escapeInlineMarkdown(text: string): string {
  // Escape the markdown the projector subset is sensitive to; keep it minimal.
  return text.replace(/([\\`*_[\]])/g, "\\$1");
}

function collectInlineSegments(node: Node, out: InlineSeg[]): void {
  node.childNodes.forEach((child) => {
    if (child.nodeType === 3 /* TEXT_NODE */) {
      const text = child.textContent ?? "";
      if (!text) return;
      const parent = child.parentElement;
      let bold = false;
      let italic = false;
      let href: string | null = null;
      if (parent && typeof getComputedStyle === "function") {
        const cs = getComputedStyle(parent);
        bold = fontWeightIsBold(cs.fontWeight);
        italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
        const a = parent.closest("a");
        href = a ? a.getAttribute("href") : null;
      }
      out.push({ text, bold, italic, href });
    } else if (child.nodeType === 1 /* ELEMENT_NODE */) {
      const el = child as HTMLElement;
      if (el.tagName === "BR") {
        out.push({ text: "\n", bold: false, italic: false, href: null });
        return;
      }
      collectInlineSegments(el, out);
    }
  });
}

function segToMarkdown(seg: InlineSeg): string {
  if (seg.text === "\n") return "\n";
  let md = escapeInlineMarkdown(seg.text);
  if (/\S/.test(seg.text)) {
    // Don't wrap pure whitespace — `** **` is not valid emphasis.
    if (seg.bold && seg.italic) md = `***${md}***`;
    else if (seg.bold) md = `**${md}**`;
    else if (seg.italic) md = `*${md}*`;
  }
  if (seg.href) md = `[${md}](${seg.href})`;
  return md;
}

/**
 * Serialize a block's inline content to the projector's markdown subset, preserving
 * bold / italic / links (emphasis detected via computed style). Used so an edit keeps
 * the block's formatting instead of flattening it to plain text. Formatting the markdown
 * subset cannot express (font size/color) is still dropped on an edited block.
 */
export function serializeInlineMarkdown(block: HTMLElement): string {
  const segs: InlineSeg[] = [];
  collectInlineSegments(block, segs);
  // Merge adjacent segments with identical formatting to avoid `**a****b**`.
  const merged: InlineSeg[] = [];
  for (const s of segs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      prev.text !== "\n" &&
      s.text !== "\n" &&
      prev.bold === s.bold &&
      prev.italic === s.italic &&
      prev.href === s.href
    ) {
      prev.text += s.text;
    } else {
      merged.push({ ...s });
    }
  }
  return merged.map(segToMarkdown).join("").trim();
}

// ─── M2: structural editing (split / merge) ─────────────────────────────────

interface AnchorRef {
  id: string;
  kind: string;
  scope: string;
  unid: string;
}

interface EditResultLite {
  success: boolean;
  created?: AnchorRef[];
  removed?: AnchorRef[];
  modified?: AnchorRef[];
  error?: { message?: string };
}

/** Plain-text character offset of the collapsed caret within `block`, or null. */
function caretOffsetIn(block: HTMLElement): number | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!block.contains(range.startContainer)) return null;
  const pre = range.cloneRange();
  pre.selectNodeContents(block);
  pre.setEnd(range.startContainer, range.startOffset);
  return pre.toString().length;
}

function placeCaretAtOffset(el: HTMLElement, offset: number): void {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel) return;
  el.focus();
  const range = document.createRange();
  let remaining = offset;
  let placed = false;
  const walk = (node: Node): void => {
    if (placed) return;
    if (node.nodeType === 3 /* TEXT_NODE */) {
      const len = node.textContent?.length ?? 0;
      if (remaining <= len) {
        range.setStart(node, remaining);
        placed = true;
      } else {
        remaining -= len;
      }
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(el);
  if (!placed) {
    range.selectNodeContents(el);
    range.collapse(false);
  } else {
    range.collapse(true);
  }
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Build the full ConvertDocxToHtmlComplete arg list (stampAnchors = last arg). */
function completeArgs(
  bytes: Uint8Array,
  cssPrefix: string,
  fabricate: boolean,
  paginated: boolean,
  scale: number,
): any[] {
  return [
    bytes, "Document", cssPrefix, fabricate, "", -1, "comment-",
    /* paginationMode */ paginated ? 1 : 0, /* paginationScale */ scale, "page-",
    false, 0, "annot-",
    /* renderFootnotesAndEndnotes */ false, /* renderHeadersAndFooters */ paginated,
    false, true, true, false, null, /* stampAnchors */ true,
  ];
}

export class DocxEditor {
  private readonly exports: DocxEditorExports;
  private readonly container: HTMLElement;
  private readonly handle: number;
  private readonly options: Required<Omit<DocxEditorOptions, "onEdit">> & Pick<DocxEditorOptions, "onEdit">;
  /** Map a block's current bare unid → its full kind:scope:unid (DocxSession anchor). */
  private readonly unidToFullId = new Map<string, string>();
  /** The element whose [data-anchor] descendants are the editable blocks (container or page container). */
  private editRoot: HTMLElement;
  private closed = false;

  private constructor(
    container: HTMLElement,
    exports: DocxEditorExports,
    handle: number,
    options: DocxEditor["options"],
  ) {
    this.container = container;
    this.exports = exports;
    this.handle = handle;
    this.options = options;
    this.editRoot = container;
  }

  /** Open a document, render it into `container`, and wire up editing. */
  static open(
    container: HTMLElement,
    bytes: Uint8Array,
    exports: DocxEditorExports,
    options: DocxEditorOptions = {},
  ): DocxEditor {
    const opts = {
      cssPrefix: options.cssPrefix ?? "docx-",
      fabricateClasses: options.fabricateClasses ?? true,
      editable: options.editable ?? true,
      paginated: options.paginated ?? false,
      scale: options.scale ?? 1,
      onEdit: options.onEdit,
    };
    const handle = exports.DocxSessionBridge.OpenSession(bytes, "");
    const editor = new DocxEditor(container, exports, handle, opts);
    editor.refreshAnchorMap();
    const fullHtml = exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(bytes, opts.cssPrefix, opts.fabricateClasses, opts.paginated, opts.scale),
    );
    if (opts.paginated) editor.mountPaginated(fullHtml);
    else editor.mountHtml(fullHtml);
    return editor;
  }

  /** Lossless DOCX bytes reflecting all edits. */
  save(): Uint8Array {
    this.assertOpen();
    return this.exports.DocxSessionBridge.Save(this.handle);
  }

  /** Release the underlying WASM session. The editor is unusable afterward. */
  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.exports.DocxSessionBridge.CloseSession(this.handle);
  }

  /** The editor's current DOM (for inspection/tests). */
  get root(): HTMLElement {
    return this.container;
  }

  // ─── internals ───────────────────────────────────────────────────────

  private assertOpen(): void {
    if (this.closed) throw new Error("DocxEditor is closed");
  }

  /** Rebuild unid → full-anchor-id from the live session projection. */
  private refreshAnchorMap(): void {
    const proj = JSON.parse(this.exports.DocxSessionBridge.Project(this.handle)) as {
      anchorIndex: Record<string, AnchorTargetLite>;
    };
    this.unidToFullId.clear();
    for (const [fullId, target] of Object.entries(proj.anchorIndex)) {
      this.unidToFullId.set(target.unid, fullId);
    }
  }

  /** Continuous (non-paginated) mount: inject the converter's styles + body, wire blocks. */
  private mountHtml(fullHtml: string): void {
    const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    this.container.innerHTML = styles + parsed.body.innerHTML;
    this.editRoot = this.container;
    if (this.options.editable) this.wireBlocks(this.container);
  }

  /** Paginated mount: flow blocks into page boxes via pagination.ts, wire the page clones. */
  private mountPaginated(fullHtml: string): void {
    paginateHtml(fullHtml, this.container, { scale: this.options.scale, cssPrefix: "page-" });
    // pagination clones blocks into pages (hidden originals stay in #pagination-staging),
    // so wire ONLY the visible page container — never the staging copies.
    const pageRoot =
      this.container.querySelector<HTMLElement>("#pagination-container") ?? this.container;
    this.editRoot = pageRoot;
    if (this.options.editable) this.wireBlocks(pageRoot);
  }

  private wireBlocks(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((el) => this.wireBlock(el));
  }

  private wireBlock(el: HTMLElement): void {
    if (!EDITABLE_TAGS.has(el.tagName)) return;
    const unid = el.getAttribute("data-anchor");
    // Only blocks the markdown projection addresses (top-level body paragraphs /
    // headings) are editable via the text path. Paragraphs inside opaque tables are
    // stamped in the HTML but not individually indexed, so they stay read-only in v1.
    if (!unid || !this.unidToFullId.has(unid)) return;
    el.setAttribute("contenteditable", "true");
    el.dataset.committedText = el.textContent ?? "";
    el.addEventListener("blur", () => this.commitBlock(el));
    el.addEventListener("keydown", (ev) => this.onKeydown(el, ev as KeyboardEvent));
  }

  /** Commit a block edit on blur: DocxSession op → re-render only this block → patch DOM. */
  private commitBlock(el: HTMLElement): void {
    if (this.closed) return;
    const unid = el.getAttribute("data-anchor");
    if (!unid) return;
    const newText = (el.textContent ?? "").trim();
    if (newText === (el.dataset.committedText ?? "").trim()) return; // no change

    const fullId = this.unidToFullId.get(unid);
    if (!fullId) return;

    // M1: serialize the block's inline content to markdown so bold/italic/links survive
    // the edit, instead of flattening to plain text.
    const markdown = serializeInlineMarkdown(el);
    const result = JSON.parse(this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, markdown)) as {
      success: boolean;
      modified?: Array<{ id: string; unid: string }>;
    };
    if (!result.success) {
      // Revert the DOM to the last committed text so the view stays in sync with truth.
      el.textContent = el.dataset.committedText ?? "";
      return;
    }

    const newAnchor = result.modified?.[0]?.id ?? fullId;
    const newUnid = result.modified?.[0]?.unid ?? unid;

    // Re-render ONLY this block from the live session and patch it in place.
    const html = this.exports.DocxSessionBridge.RenderBlockHtml(
      this.handle,
      newAnchor,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) !== 0x7b /* not an error object */) {
      const fresh = new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
      if (fresh) {
        el.replaceWith(fresh);
        // Keep the anchor map current and re-wire the replacement.
        this.unidToFullId.delete(unid);
        this.unidToFullId.set(newUnid, newAnchor);
        this.wireBlock(fresh);
      }
    }

    this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
  }

  // ─── M2: structural editing ──────────────────────────────────────────

  private onKeydown(el: HTMLElement, ev: KeyboardEvent): void {
    if (this.closed) return;
    if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      this.splitAtCaret(el);
    } else if (ev.key === "Backspace") {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (sel && sel.isCollapsed && caretOffsetIn(el) === 0) {
        const prev = this.previousEditable(el);
        if (prev) {
          ev.preventDefault();
          this.mergeWithPrevious(prev, el);
        }
      }
    }
  }

  /** Enter: split the block at the caret into two paragraphs. */
  private splitAtCaret(el: HTMLElement): void {
    const offset = caretOffsetIn(el);
    const unid = el.getAttribute("data-anchor");
    if (offset == null || !unid) return;
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;

    fullId = this.syncBlock(el, fullId); // flush any uncommitted typing first
    const res = this.parseEdit(this.exports.DocxSessionBridge.SplitParagraph(this.handle, fullId, offset));
    if (!res.success) return;
    const first = res.modified?.[0];
    const second = res.created?.[0];
    if (!first || !second) return;

    const firstEl = this.renderInto(first.id);
    const secondEl = this.renderInto(second.id);
    if (!firstEl || !secondEl) return;

    el.replaceWith(firstEl);
    firstEl.after(secondEl);
    this.unidToFullId.delete(unid);
    this.unidToFullId.set(first.unid, first.id);
    this.unidToFullId.set(second.unid, second.id);
    this.wireBlock(firstEl);
    this.wireBlock(secondEl);
    placeCaretAtOffset(secondEl, 0);
    this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
  }

  /** Backspace at block start: merge this block into the previous one. */
  private mergeWithPrevious(prev: HTMLElement, el: HTMLElement): void {
    const prevUnid = prev.getAttribute("data-anchor");
    const thisUnid = el.getAttribute("data-anchor");
    if (!prevUnid || !thisUnid) return;
    let prevId = this.unidToFullId.get(prevUnid);
    let thisId = this.unidToFullId.get(thisUnid);
    if (!prevId || !thisId) return;

    prevId = this.syncBlock(prev, prevId);
    thisId = this.syncBlock(el, thisId);
    const caret = (prev.textContent ?? "").length; // merge boundary

    const res = this.parseEdit(this.exports.DocxSessionBridge.MergeParagraphs(this.handle, prevId, thisId));
    if (!res.success) return;
    const merged = res.modified?.[0];
    if (!merged) return;

    const mergedEl = this.renderInto(merged.id);
    if (!mergedEl) return;

    prev.replaceWith(mergedEl);
    el.remove();
    this.unidToFullId.delete(prevUnid);
    this.unidToFullId.delete(thisUnid);
    this.unidToFullId.set(merged.unid, merged.id);
    this.wireBlock(mergedEl);
    placeCaretAtOffset(mergedEl, caret);
    this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
  }

  /** Flush a block's current (uncommitted) text to the session; returns the live full id. */
  private syncBlock(el: HTMLElement, fullId: string): string {
    const cur = (el.textContent ?? "").trim();
    if (cur === (el.dataset.committedText ?? "").trim()) return fullId; // unchanged
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, serializeInlineMarkdown(el)),
    );
    return res.success ? res.modified?.[0]?.id ?? fullId : fullId;
  }

  /** Render a block by anchor and parse it into a detached element (null on error). */
  private renderInto(anchorId: string): HTMLElement | null {
    const html = this.exports.DocxSessionBridge.RenderBlockHtml(
      this.handle,
      anchorId,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) === 0x7b /* error object */) return null;
    return new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
  }

  /** The editable block immediately before `el` in document order, or null. */
  private previousEditable(el: HTMLElement): HTMLElement | null {
    const all = Array.from(
      this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    const i = all.indexOf(el);
    return i > 0 ? all[i - 1] : null;
  }

  private parseEdit(json: string): EditResultLite {
    try {
      return JSON.parse(json) as EditResultLite;
    } catch {
      return { success: false };
    }
  }
}
