#nullable enable
using Docxodus;
using Xunit;
namespace Docxodus.Tests;
public class DocxDiffConsolidateApiTests
{
    [Fact]
    public void Settings_default_policy_is_base_wins()
    {
        var s = new DocxDiffConsolidateSettings();
        Assert.Equal(ConflictResolution.BaseWins, s.ConflictResolution);
        Assert.NotNull(s.Diff);
    }

    [Fact]
    public void Reviewer_holds_document_and_author()
    {
        var r = new DocxDiffReviewer { Author = "Bob" };
        Assert.Equal("Bob", r.Author);
    }
}
