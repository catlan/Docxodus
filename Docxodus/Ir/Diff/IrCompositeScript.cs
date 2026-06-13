#nullable enable

using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// One token op tagged with the reviewer who authored it, used when composing multi-reviewer
/// modified paragraphs. <see cref="SourceReviewer"/> indexes the caller's reviewers list;
/// -1 means base-sourced (Equal/Delete spans).
/// </summary>
internal sealed record IrAuthoredTokenOp(IrTokenOp Op, string Author, int SourceReviewer);

/// <summary>
/// Maps one contributing reviewer (<see cref="Reviewer"/>, indexing the caller's reviewers list) to the
/// <c>kind:scope:unid</c> RIGHT anchor of that reviewer's edited paragraph for a composed
/// <see cref="IrCompositeOp"/>. The composite renderer needs each contributing reviewer's OWN right
/// paragraph to source inserted runs: an <see cref="IrAuthoredTokenOp"/>'s Insert RightStart/RightEnd
/// index THAT reviewer's right-token list. <see cref="IrCompositeOp.SourceRightAnchors"/> carries one
/// entry per contributing reviewer so the renderer can resolve each right paragraph independently.
/// </summary>
internal sealed record IrSourceRightAnchor(int Reviewer, string Anchor);

/// <summary>
/// An edit op tagged with its contributing reviewer. For a composed multi-reviewer Modify,
/// <see cref="Op"/>'s <c>TokenDiff</c> is the MERGED diff (apply/json truth) and
/// <see cref="AuthoredTokens"/> carries per-span authorship for the renderer; for all
/// single-source ops <see cref="AuthoredTokens"/> is null and <see cref="Author"/>/<see cref="SourceReviewer"/>
/// apply. <see cref="SourceReviewer"/> -1 = base-sourced.
/// <para><see cref="ConflictId"/> is non-null when the op is the winner-representative of a conflict
/// (the losing competitors are recorded in <see cref="IrCompositeScript.Conflicts"/>).</para>
/// <para><see cref="SourceRightAnchors"/> is non-null ONLY on a composed multi-reviewer Modify (alongside
/// <see cref="AuthoredTokens"/>): it carries, per contributing reviewer, the right paragraph anchor the
/// renderer must resolve to source that reviewer's inserted runs (their <see cref="IrAuthoredTokenOp"/>
/// Insert spans index that reviewer's right-token list). For all single-source ops it is null and
/// <see cref="Op"/>'s own <c>RightAnchor</c> suffices. Additive/optional — absent from older scripts and
/// from the JSON wire shape, so existing tests/serialization are unaffected.</para>
/// </summary>
internal sealed record IrCompositeOp(
    IrEditOp Op,
    string Author,
    int SourceReviewer,
    IrNodeList<IrAuthoredTokenOp>? AuthoredTokens = null,
    int? ConflictId = null,
    IrNodeList<IrSourceRightAnchor>? SourceRightAnchors = null);

/// <summary>
/// One reviewer's competing result for a conflicted span. <see cref="Author"/> is the reviewer
/// name; <see cref="ResultText"/> is the flat text the reviewer's edit would have produced.
/// </summary>
internal sealed record IrConflictCompetitor(string Author, string ResultText);

/// <summary>
/// A base span (<see cref="TokenStart"/>..<see cref="TokenEnd"/>, [x,x) for a block-level conflict)
/// edited DIFFERENTLY by 2+ reviewers. <see cref="Id"/> links back to the <see cref="IrCompositeOp.ConflictId"/>
/// of the winning op in <see cref="IrCompositeScript.Operations"/>. <see cref="AppliedPolicy"/> records
/// which <see cref="ConflictResolution"/> was used; <see cref="Competitors"/> lists every competing edit
/// (including the one that won, for audit purposes).
/// </summary>
internal sealed record IrConflict(
    int Id,
    string BaseAnchor,
    int TokenStart,
    int TokenEnd,
    Docxodus.ConflictResolution AppliedPolicy,
    IrNodeList<IrConflictCompetitor> Competitors);

/// <summary>
/// The composite diff-as-data product: base-anchor-ordered authored ops + the conflict list.
/// <see cref="NoteOps"/> carries per-note block diffs (footnotes/endnotes) when present,
/// mirroring the <see cref="IrEditScript.NoteOps"/> convention.
/// </summary>
internal sealed record IrCompositeScript(
    IrNodeList<IrCompositeOp> Operations,
    IrNodeList<IrConflict> Conflicts,
    IrNodeList<IrNoteDiff>? NoteOps = null);
