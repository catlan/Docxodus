#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// N-way merger that consolidates several pairwise <see cref="IrEditScript"/>s — each computed
/// against the SAME base document, so they share one base anchor space — into a single
/// <see cref="IrCompositeScript"/>.
/// <para>The merge walks the base document in block order. For each base block it gathers the
/// reviewers' ops anchored there (via <see cref="GroupByBaseAnchor"/>) and dispatches to
/// <c>MergeOneBaseBlock</c>; right-only inserts are slotted relative to the preceding base anchor
/// (via <c>GroupInsertsByPrecedingAnchor</c> + <c>EmitInsertsAt</c>). Conflicts — a base block
/// edited differently by 2+ reviewers — are resolved by the supplied
/// <see cref="Docxodus.ConflictResolution"/> policy and recorded in
/// <see cref="IrCompositeScript.Conflicts"/>.</para>
/// <para>See <c>docs/architecture/ir_diff_engine.md</c> for the diff pipeline this composes over.</para>
/// </summary>
internal static class IrCompositeMerger
{
    public static IrCompositeScript Merge(
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings)
    {
        var scripts = reviewers.Select(r => IrEditScriptBuilder.Build(baseIr, r.Ir, settings)).ToList();
        var baseOrder = BaseBlockAnchors(baseIr);
        var byBase = GroupByBaseAnchor(scripts);
        var insertsAfter = GroupInsertsByPrecedingAnchor(scripts, baseOrder);

        var ops = new List<IrCompositeOp>();
        var conflicts = new List<IrConflict>();
        var nextConflictId = 1;
        EmitInsertsAt(insertsAfter, "", reviewers, ops);
        foreach (var anchor in baseOrder)
        {
            MergeOneBaseBlock(anchor, byBase, reviewers, baseIr, policy, settings, ops, conflicts, ref nextConflictId);
            EmitInsertsAt(insertsAfter, anchor, reviewers, ops);
        }
        return new IrCompositeScript(IrNodeList.From(ops), IrNodeList.From(conflicts));
    }

    /// <summary>
    /// Index every op that carries a non-null <see cref="IrEditOp.LeftAnchor"/> (i.e. it touches a
    /// base block) by that anchor string, tagged with its reviewer index in <paramref name="scripts"/>.
    /// </summary>
    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupByBaseAnchor(
        IReadOnlyList<IrEditScript> scripts)
    {
        var byBase = new Dictionary<string, List<(int Reviewer, IrEditOp Op)>>();
        for (int reviewer = 0; reviewer < scripts.Count; reviewer++)
        {
            foreach (var op in scripts[reviewer].Operations)
            {
                if (op.LeftAnchor is not { } anchor)
                    continue;
                if (!byBase.TryGetValue(anchor, out var list))
                    byBase[anchor] = list = new List<(int, IrEditOp)>();
                list.Add((reviewer, op));
            }
        }
        return byBase;
    }

    /// <summary>Base body blocks in document order, as <c>kind:scope:unid</c> anchor strings.</summary>
    internal static List<string> BaseBlockAnchors(IrDocument baseIr) =>
        baseIr.Body.Blocks.Select(b => b.Anchor.ToString()).ToList();

    // ---- Task 1.3: exact sub-block (token-span) composition ----

    /// <summary>
    /// Compose N reviewers' per-base-paragraph token diffs — all expressed over the SAME base token
    /// stream <c>[0, baseTokenCount)</c> — into a single authored token-op list plus any conflicts
    /// (spec §6 step 4). Walks base positions <c>pos = 0..baseTokenCount</c> inclusive; at each pos it
    /// (1) resolves the reviewers' INSERTs anchored at pos (consensus when ≤1 distinct inserted text,
    /// otherwise a conflict resolved by <paramref name="policy"/>), then (2) for the base token
    /// <c>[pos, pos+1)</c> resolves the reviewers' DELETEs (consensus when 0/1 deleter or all deleters
    /// delete identically, otherwise a conflict). Conflict ids ascend from <paramref name="conflictIdSeed"/>.
    /// </summary>
    /// <remarks>
    /// INVARIANT (span totality): the emitted ops' non-insert left spans tile <c>[0, baseTokenCount)</c>
    /// exactly once — every base token becomes exactly one Equal or one Delete; inserts carry empty
    /// left spans. Verified by <see cref="AssertTilesBase"/> under <c>Debug.Assert</c>.
    /// </remarks>
    internal static List<IrAuthoredTokenOp> ComposeTokenSpans(
        int baseTokenCount,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, int conflictIdSeed,
        out List<IrConflict> conflicts)
    {
        var result = new List<IrAuthoredTokenOp>();
        conflicts = new List<IrConflict>();
        int nextConflictId = conflictIdSeed;

        for (int pos = 0; pos <= baseTokenCount; pos++)
        {
            ComposeInsertsAt(pos, reviewerDiffs, policy, baseAnchor, ref nextConflictId, result, conflicts);
            if (pos < baseTokenCount)
                ComposeBaseTokenAt(pos, reviewerDiffs, policy, baseAnchor, ref nextConflictId, result, conflicts);
        }

        System.Diagnostics.Debug.Assert(AssertTilesBase(result, baseTokenCount),
            "ComposeTokenSpans: emitted left spans must tile [0, baseTokenCount) exactly once.");
        return result;
    }

    /// <summary>Step 1: resolve every reviewer INSERT anchored at base position <paramref name="pos"/>.</summary>
    private static void ComposeInsertsAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each entry: the inserting reviewer, in reviewer (list) order, with the inserted text.
        var inserters = new List<(int ReviewerIndex, string Author, IrTokenOp Op, string Text)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            foreach (var op in rd.Diff.Ops)
            {
                if (op.Kind == IrTokenOpKind.Insert && op.LeftStart == pos)
                    inserters.Add((rd.Reviewer, rd.Author, op, RightText(op, rd.RightTokens)));
            }
        }
        if (inserters.Count == 0)
            return;

        // Group by inserted right text, preserving first-appearance order (reviewer order).
        var distinctTexts = new List<string>();
        foreach (var ins in inserters)
            if (!distinctTexts.Contains(ins.Text))
                distinctTexts.Add(ins.Text);

        if (distinctTexts.Count <= 1)
        {
            // Consensus (single reviewer, or multiple reviewers inserting the SAME text): one authored
            // Insert per distinct text; the first reviewer in that group is the author/source.
            foreach (var text in distinctTexts)
            {
                var first = inserters.First(x => x.Text == text);
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
            }
            return;
        }

        // ≥2 distinct text groups → CONFLICT: one competitor per inserting reviewer.
        var competitors = inserters
            .Select(x => new IrConflictCompetitor(x.Author, x.Text))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Emit NOTHING for the inserts at pos.
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                var first = inserters.OrderBy(x => x.ReviewerIndex).First();
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
                foreach (var ins in inserters.OrderBy(x => x.ReviewerIndex))
                    result.Add(new IrAuthoredTokenOp(ins.Op, ins.Author, ins.ReviewerIndex));
                break;
        }
    }

    /// <summary>Step 2: resolve the fate of base token <c>[pos, pos+1)</c> across reviewer DELETEs.</summary>
    private static void ComposeBaseTokenAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each deleter, in reviewer order, with its replacement text ("" for a pure delete). A
        // replacement is an Insert anchored at the SAME base pos in the same reviewer's diff.
        var deleters = new List<(int ReviewerIndex, string Author, IrTokenOp DeleteOp, string Replacement)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            IrTokenOp? delOp = null;
            foreach (var op in rd.Diff.Ops)
            {
                if (op.Kind == IrTokenOpKind.Delete && op.LeftStart <= pos && pos < op.LeftEnd)
                {
                    delOp = op;
                    break;
                }
            }
            if (delOp is null)
                continue;
            var replacement = ReplacementTextAt(pos, rd);
            deleters.Add((rd.Reviewer, rd.Author, delOp, replacement));
        }

        if (deleters.Count == 0)
        {
            // Base survives.
            result.Add(BaseEqual(pos));
            return;
        }

        // "Identical" deletion: all deleters have byte-equal replacement text (all pure deletes, or
        // all delete-and-replace with the same inserted text). Pure-delete vs delete-with-replacement
        // are NOT identical (one has "" replacement, the other does not).
        bool allIdentical = deleters.All(d => d.Replacement == deleters[0].Replacement);

        if (deleters.Count == 1 || allIdentical)
        {
            // Authored to the (first) deleter — delete the base token once.
            var first = deleters[0];
            result.Add(new IrAuthoredTokenOp(first.DeleteOp, first.Author, first.ReviewerIndex));
            return;
        }

        // ≥2 deleters that are NOT identical → CONFLICT.
        var competitors = deleters
            .Select(d => new IrConflictCompetitor(d.Author, d.Replacement))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos + 1, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Base survives.
                result.Add(BaseEqual(pos));
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                var first = deleters.OrderBy(d => d.ReviewerIndex).First();
                result.Add(new IrAuthoredTokenOp(first.DeleteOp, first.Author, first.ReviewerIndex));
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
            {
                // A base token can only be deleted once, so stacking multiple Deletes of the same base
                // token is degenerate: emit just the first deleter's Delete. The competing REPLACEMENT
                // inserts are emitted (stacked) by the insert path (ComposeInsertsAt) at the same pos.
                var first = deleters.OrderBy(d => d.ReviewerIndex).First();
                result.Add(new IrAuthoredTokenOp(first.DeleteOp, first.Author, first.ReviewerIndex));
                break;
            }
        }
    }

    /// <summary>A base-sourced Equal op covering base token <paramref name="pos"/> (Author "", reviewer -1).</summary>
    private static IrAuthoredTokenOp BaseEqual(int pos) =>
        new(new IrTokenOp(IrTokenOpKind.Equal, pos, pos + 1, pos, pos + 1), "", -1);

    /// <summary>
    /// The replacement text a reviewer attaches to deleting base token <paramref name="pos"/>: the
    /// concatenated raw text of its Insert anchored at the same base pos, or "" for a pure delete.
    /// </summary>
    private static string ReplacementTextAt(
        int pos, (int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens) rd)
    {
        foreach (var op in rd.Diff.Ops)
            if (op.Kind == IrTokenOpKind.Insert && op.LeftStart == pos)
                return RightText(op, rd.RightTokens);
        return "";
    }

    /// <summary>Concatenate the raw <see cref="IrDiffToken.Text"/> of <paramref name="op"/>'s right span.</summary>
    private static string RightText(IrTokenOp op, IReadOnlyList<IrDiffToken> rightTokens)
    {
        if (op.RightLength == 0)
            return "";
        var sb = new System.Text.StringBuilder();
        for (int i = op.RightStart; i < op.RightEnd; i++)
            sb.Append(rightTokens[i].Text);
        return sb.ToString();
    }

    /// <summary>
    /// Debug-only totality check: the concatenation of non-insert (Equal/Delete) left spans, in op
    /// order, equals <c>[0, baseTokenCount)</c> exactly once ascending.
    /// </summary>
    private static bool AssertTilesBase(List<IrAuthoredTokenOp> ops, int baseTokenCount)
    {
        int expected = 0;
        foreach (var authored in ops)
        {
            var op = authored.Op;
            if (op.Kind == IrTokenOpKind.Insert)
                continue;
            if (op.LeftStart != expected)
                return false;
            expected = op.LeftEnd;
        }
        return expected == baseTokenCount;
    }

    // ---- Filled by Task 1.4 ----

    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupInsertsByPrecedingAnchor(
        IReadOnlyList<IrEditScript> scripts, List<string> baseOrder) => throw new System.NotImplementedException();

    private static void EmitInsertsAt(
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> insertsAfter,
        string anchor,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        List<IrCompositeOp> ops) => throw new System.NotImplementedException();

    private static void MergeOneBaseBlock(
        string anchor,
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> byBase,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings,
        List<IrCompositeOp> ops,
        List<IrConflict> conflicts,
        ref int nextConflictId) => throw new System.NotImplementedException();
}
