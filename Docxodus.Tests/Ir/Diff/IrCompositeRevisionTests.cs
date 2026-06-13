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
}
