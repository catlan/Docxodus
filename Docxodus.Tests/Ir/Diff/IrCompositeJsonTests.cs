#nullable enable
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeJsonTests
{
    [Fact]
    public void Consolidated_edit_script_json_has_author_and_conflicts()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox"); var r2 = Docs.Para("the FAST brown fox");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"author\"", json);
        Assert.Contains("\"conflicts\"", json);
        using var _ = System.Text.Json.JsonDocument.Parse(json); // valid JSON
    }

    [Fact]
    public void Consolidated_edit_script_json_carries_source_reviewer_and_anchors()
    {
        var b = Docs.Para("alpha one", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "gamma three");
        var r2 = Docs.Para("alpha one", "gamma three EDITED");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"sourceReviewer\"", json);
        Assert.Contains("\"operations\"", json);
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("conflicts", out _));
        Assert.True(doc.RootElement.TryGetProperty("operations", out _));
    }

    [Fact]
    public void Zero_reviewers_json_is_empty_operations_and_conflicts()
    {
        var b = Docs.Para("alpha one");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, System.Array.Empty<DocxDiffReviewer>());
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.Empty(doc.RootElement.GetProperty("operations").EnumerateArray());
        Assert.Empty(doc.RootElement.GetProperty("conflicts").EnumerateArray());
    }
}
