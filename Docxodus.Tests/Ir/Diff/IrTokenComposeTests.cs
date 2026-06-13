#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.4 Task 1.3 tests for <see cref="IrCompositeMerger.ComposeTokenSpans"/>: the EXACT sub-block
/// merge that composes N reviewers' per-base-paragraph token diffs (all over the SAME base token
/// stream) into one authored token-op list plus the conflict list.
/// </summary>
/// <remarks>
/// Reviewer right-token lists are built as synthetic <see cref="IrDiffToken"/> Word tokens whose
/// raw <c>Text</c> is the inserted/replacement string. Inserts are modeled as a single Word token
/// per distinct text so <c>RightStart..RightEnd</c> is a one-token span; this matches how the merger
/// concatenates raw text to decide consensus vs conflict.
/// </remarks>
public class IrTokenComposeTests
{
    private static readonly IrRunFormat Plain = new() { Bold = false, UnmodeledDigest = default };

    /// <summary>One Word token whose Text/MatchKey are <paramref name="text"/>.</summary>
    private static IrDiffToken Tok(string text) =>
        new(IrDiffTokenKind.Word, text, text, 0, text.Length, Plain);

    /// <summary>A right-token list of single Word tokens, one per supplied string.</summary>
    private static List<IrDiffToken> Rights(params string[] texts) =>
        texts.Select(Tok).ToList();

    /// <summary>An Insert op anchored at base position <paramref name="pos"/> spanning right tokens
    /// <paramref name="rStart"/>..<paramref name="rEnd"/>.</summary>
    private static IrTokenOp Ins(int pos, int rStart, int rEnd) =>
        new(IrTokenOpKind.Insert, pos, pos, rStart, rEnd);

    /// <summary>A Delete op of the base token at <paramref name="pos"/> (right span empty at
    /// <paramref name="rAnchor"/>).</summary>
    private static IrTokenOp Del(int pos, int rAnchor) =>
        new(IrTokenOpKind.Delete, pos, pos + 1, rAnchor, rAnchor);

    /// <summary>An Equal op covering base token <paramref name="pos"/>.</summary>
    private static IrTokenOp Eq(int pos, int rPos) =>
        new(IrTokenOpKind.Equal, pos, pos + 1, rPos, rPos + 1);

    private static (int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>) Reviewer(
        int idx, string author, IReadOnlyList<IrDiffToken> rights, params IrTokenOp[] ops) =>
        (idx, author, new IrTokenDiff(IrNodeList.From(ops)), rights);

    [Fact]
    public void Disjoint_token_inserts_compose_without_conflict()
    {
        // base count 5; R1 inserts "AAA" at pos 1; R2 inserts "BBB" at pos 3.
        var r1Rights = Rights("AAA");
        var r2Rights = Rights("BBB");
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", r1Rights, Ins(1, 0, 1)),
            Reviewer(1, "Fred", r2Rights, Ins(3, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Equal(2, inserts.Count);
        var authors = inserts.Select(o => o.Author).ToHashSet();
        Assert.Contains("Bob", authors);
        Assert.Contains("Fred", authors);
    }

    [Fact]
    public void Same_text_insert_by_two_reviewers_is_consensus()
    {
        var r1Rights = Rights("SAME");
        var r2Rights = Rights("SAME");
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", r1Rights, Ins(2, 0, 1)),
            Reviewer(1, "Fred", r2Rights, Ins(2, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Single(inserts);
        Assert.Equal("Bob", inserts[0].Author); // first reviewer in the group is the author
    }

    [Fact]
    public void Different_text_insert_at_same_anchor_conflicts()
    {
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights("X"), Ins(2, 0, 1)),
            Reviewer(1, "Fred", Rights("Y"), Ins(2, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 7, out var conflicts);

        Assert.Single(conflicts);
        Assert.Equal(7, conflicts[0].Id);
        Assert.Equal(2, conflicts[0].TokenStart);
        Assert.Equal(2, conflicts[0].TokenEnd);
        Assert.Equal(2, conflicts[0].Competitors.Count);
        // BaseWins → no insert at pos 2.
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    [Fact]
    public void Overlapping_replacement_conflicts_basewins_keeps_base()
    {
        // R1: delete base token[1], insert "X" at pos1. R2: delete base token[1], insert "Y".
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights("X"), Del(1, 0), Ins(1, 0, 1)),
            Reviewer(1, "Fred", Rights("Y"), Del(1, 0), Ins(1, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        // Two conflicts expected: the differing-text insert at pos 1, and the
        // delete-with-different-replacement at base token 1. With BaseWins the base survives.
        Assert.NotEmpty(conflicts);
        // Base token 1 survives as Equal (no Delete at LeftStart 1).
        Assert.DoesNotContain(ops, o => o.Op.Kind == IrTokenOpKind.Delete && o.Op.LeftStart == 1);
        Assert.Contains(ops, o => o.Op.Kind == IrTokenOpKind.Equal && o.Op.LeftStart == 1);
        // BaseWins → no insert emitted at the conflicted anchor.
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    [Fact]
    public void Pure_delete_consensus()
    {
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights(), Del(1, 0)),
            Reviewer(1, "Fred", Rights(), Del(1, 0)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        var deletes = ops.Where(o => o.Op.Kind == IrTokenOpKind.Delete).ToList();
        Assert.Single(deletes);
        Assert.Equal(1, deletes[0].Op.LeftStart);
        Assert.Equal("Bob", deletes[0].Author);
    }

    [Fact]
    public void FirstReviewerWins_and_StackAll_emit_expected()
    {
        List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)> Build() => new()
        {
            Reviewer(0, "Bob", Rights("X"), Ins(2, 0, 1)),
            Reviewer(1, "Fred", Rights("Y"), Ins(2, 0, 1)),
        };

        var firstWins = IrCompositeMerger.ComposeTokenSpans(
            5, Build(), ConflictResolution.FirstReviewerWins, "p:body:base", 1, out var c1);
        Assert.Single(c1);
        var fwInserts = firstWins.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Single(fwInserts);
        Assert.Equal("Bob", fwInserts[0].Author);

        var stackAll = IrCompositeMerger.ComposeTokenSpans(
            5, Build(), ConflictResolution.StackAll, "p:body:base", 1, out var c2);
        Assert.Single(c2);
        var saInserts = stackAll.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Equal(2, saInserts.Count);
        Assert.Equal(new[] { "Bob", "Fred" }, saInserts.Select(o => o.Author).ToArray());
    }
}
