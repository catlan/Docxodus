#nullable enable
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeMergerTests
{
    internal static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    [Fact]
    public void GroupByBaseAnchor_colocates_per_reviewer_ops()
    {
        var baseDoc = Docs.Para("alpha one", "beta two");
        var r1 = Docs.Para("alpha one EDITED", "beta two");
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var s1 = IrEditScriptBuilder.Build(baseIr, IrReader.Read(r1, ReadOpts), diff);
        var grouped = IrCompositeMerger.GroupByBaseAnchor(new[] { s1 });
        Assert.Contains(grouped.Values, list => list.Any(x => x.Op.Kind == IrEditOpKind.ModifyBlock));
    }

    // Helper reused by later tasks: merge base + reviewers into an IrCompositeScript.
    internal static IrCompositeScript MergeOf(WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
        => MergeOf(ConflictResolution.BaseWins, baseDoc, reviewers);

    internal static IrCompositeScript MergeOf(ConflictResolution policy, WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revs = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        return IrCompositeMerger.Merge(baseIr, revs, policy, diff);
    }
}
