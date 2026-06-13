#nullable enable
using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeRevisionTests
{
    [Fact]
    public void Consolidated_revisions_carry_per_reviewer_author()
    {
        var b = Docs.Para("alpha one", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "gamma three");
        var r2 = Docs.Para("alpha one", "gamma three EDITED");
        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains(revs, x => x.Author == "Bob");
        Assert.Contains(revs, x => x.Author == "Fred");
    }

    [Fact]
    public void Conflicted_revision_carries_conflict_id()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox"); var r2 = Docs.Para("the FAST brown fox");
        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}},
            new DocxDiffConsolidateSettings{ConflictResolution=ConflictResolution.StackAll});
        Assert.Contains(revs, x => x.ConflictId != null);
    }

    [Fact]
    public void Zero_reviewers_yields_no_revisions()
    {
        var b = Docs.Para("alpha one");
        var revs = DocxDiff.GetConsolidatedRevisions(b, System.Array.Empty<DocxDiffReviewer>());
        Assert.Empty(revs);
    }

    [Fact]
    public void Insertion_and_deletion_attributed_to_authoring_reviewer()
    {
        var b = Docs.Para("keep this");
        var r1 = Docs.Para("keep this added");   // Bob inserts "added"
        var r2 = Docs.Para("this");              // Fred deletes "keep"
        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains(revs, x => x.Type == DocxDiffRevisionType.Inserted && x.Author == "Bob");
        Assert.Contains(revs, x => x.Type == DocxDiffRevisionType.Deleted && x.Author == "Fred");
    }

    [Fact]
    public void Replace_conflict_stacks_link_both_delete_and_insert_to_conflict()
    {
        // Two reviewers REPLACE the same base word with different text → a delete/replace conflict on
        // [1,2). The replacement inserts ("SLOW"/"FAST") are anchored at the delete's LeftEnd (token 2,
        // == conflict TokenEnd), so the conflict id must reach BOTH the delete revisions AND the inserts.
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox"); var r2 = Docs.Para("the FAST brown fox");
        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}},
            new DocxDiffConsolidateSettings{ConflictResolution=ConflictResolution.StackAll});

        var slow = Assert.Single(revs, x => x.Text.Contains("SLOW") && x.Type == DocxDiffRevisionType.Inserted);
        var fast = Assert.Single(revs, x => x.Text.Contains("FAST") && x.Type == DocxDiffRevisionType.Inserted);
        Assert.NotNull(slow.ConflictId);
        Assert.NotNull(fast.ConflictId);
    }

    [Fact]
    public void Insert_vs_insert_conflict_links_conflict_id()
    {
        // Two reviewers insert DIFFERENT text at the SAME base position with no surrounding delete →
        // a zero-width insert-vs-insert conflict [pos,pos). The inserted-text revisions must carry it.
        var b = Docs.Para("alpha omega");
        var r1 = Docs.Para("alpha INSERTEDBOB omega");
        var r2 = Docs.Para("alpha INSERTEDFRED omega");
        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}},
            new DocxDiffConsolidateSettings{ConflictResolution=ConflictResolution.StackAll});

        // Confirm the merger actually produced a zero-width insert conflict for this input.
        var conflicts = DocxDiff.GetConflicts(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}},
            new DocxDiffConsolidateSettings{ConflictResolution=ConflictResolution.StackAll});
        Assert.Contains(conflicts, c => c.TokenStart == c.TokenEnd);

        var bob = Assert.Single(revs, x => x.Text.Contains("INSERTEDBOB") && x.Type == DocxDiffRevisionType.Inserted);
        var fred = Assert.Single(revs, x => x.Text.Contains("INSERTEDFRED") && x.Type == DocxDiffRevisionType.Inserted);
        Assert.NotNull(bob.ConflictId);
        Assert.NotNull(fred.ConflictId);
    }
}
