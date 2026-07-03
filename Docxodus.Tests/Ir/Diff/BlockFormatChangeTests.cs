#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Phase 0 characterization pins for the block-format-change family
/// (spec: docs/superpowers/specs/2026-07-03-diff-block-format-changes-design.md §1).
/// Each region pins what pairwise Compare/GetRevisions does TODAY for a property-only change with
/// identical text, at the soundness level (is the change visible? tracked? does accept ≡ right and
/// reject ≡ left hold at the property level?). As each implementation phase lands, the corresponding
/// pins are FLIPPED in place to the new tracked behavior — never deleted.
/// </summary>
public class BlockFormatChangeTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly DocxDiffSettings ModeledOnly = new();
    private static readonly DocxDiffSettings Full = new() { FormatComparison = DocxDiffFormatComparison.Full };

    private static XElement BodyOf(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        using var s = wDoc.MainDocumentPart!.GetStream();
        return XElement.Load(s).Element(W + "body")!;
    }

    // ------------------------------------------------------------------ pPr-only (w:jc)

    private static readonly WmlDocument PPrLeft = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>");

    private static readonly WmlDocument PPrRight = IrTestDocuments.FromBodyXml(
        "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Same text here.</w:t></w:r></w:p>");

    [Fact]
    public void PPrOnly_ModeledOnly_is_invisible_and_untracked_right_apply_today()
    {
        // Pin: BlockSignature is run-token-only, so a modeled pPr-only change classifies Unchanged.
        Assert.Empty(DocxDiff.GetRevisions(PPrLeft, PPrRight, ModeledOnly));

        var result = DocxDiff.Compare(PPrLeft, PPrRight, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "pPrChange"));                                  // untracked
        Assert.Equal("center", (string?)body.Descendants(W + "jc").Single().Attribute(W + "val")); // right applied

        // The soundness pin: reject does NOT restore the left pPr (right jc survives rejection).
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("center", (string?)BodyOf(rejected).Descendants(W + "jc").Single().Attribute(W + "val"));
    }

    [Fact]
    public void PPrOnly_Full_is_visible_but_untracked_today()
    {
        // Pin: under Full the fingerprint flips → FormatOnly → ONE whole-block FormatChanged revision
        // with EMPTY details (the delta is not describable at run-token grain).
        var revisions = DocxDiff.GetRevisions(PPrLeft, PPrRight, Full);
        var rev = Assert.Single(revisions);
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        Assert.NotNull(rev.FormatChange);
        Assert.Empty(rev.FormatChange!.ChangedPropertyNames);

        // But the markup is still an untracked right-apply: no pPrChange, reject ≠ left.
        var result = DocxDiff.Compare(PPrLeft, PPrRight, Full);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "pPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("center", (string?)BodyOf(rejected).Descendants(W + "jc").Single().Attribute(W + "val"));
    }

    // ------------------------------------------------------------------ trPr-only (w:trHeight)

    private static string Table(string trPr, string tcPr, string tblPr = "<w:tblW w:w=\"0\" w:type=\"auto\"/>",
                                string grid = "<w:gridCol w:w=\"4000\"/>") =>
        $"<w:tbl><w:tblPr>{tblPr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>" +
        $"<w:tr>{trPr}<w:tc><w:tcPr>{tcPr}</w:tcPr><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>After.</w:t></w:r></w:p>";

    [Fact]
    public void TrPrOnly_is_untracked_right_apply_today()
    {
        var left = IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"400\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"800\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));

        // Pin: FormatOnly (fingerprint-grade for non-paragraphs) but nothing describable → no revisions.
        Assert.Empty(DocxDiff.GetRevisions(left, right, ModeledOnly));

        // Pin: non-paragraph FormatOnly falls through to a verbatim right emit — no trPrChange, right trHeight.
        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "trPrChange"));
        Assert.Equal("800", (string?)body.Descendants(W + "trHeight").Single().Attribute(W + "val"));

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("800", (string?)BodyOf(rejected).Descendants(W + "trHeight").Single().Attribute(W + "val")); // reject ≠ left
    }

    // ------------------------------------------------------------------ tblPr-only (w:tblBorders)

    [Fact]
    public void TblPrOnly_is_untracked_right_apply_today()
    {
        var borders = "<w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders>";
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>", tblPr: borders));

        Assert.Empty(DocxDiff.GetRevisions(left, right, ModeledOnly));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "tblPrChange"));
        Assert.Single(body.Descendants(W + "tblBorders"));                                 // right applied

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Single(BodyOf(rejected).Descendants(W + "tblBorders"));                     // reject ≠ left
    }

    // ------------------------------------------------------------------ tblGrid-only (column width)

    [Fact]
    public void TblGridOnly_is_untracked_right_apply_today()
    {
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>", grid: "<w:gridCol w:w=\"6000\"/>"));

        Assert.Empty(DocxDiff.GetRevisions(left, right, ModeledOnly));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "tblGridChange"));
        Assert.Equal("6000", (string?)body.Descendants(W + "gridCol").Single().Attribute(W + "w"));

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("6000", (string?)BodyOf(rejected).Descendants(W + "gridCol").Single().Attribute(W + "w"));
    }

    // ------------------------------------------------------------------ tcPr-only (w:shd)

    [Fact]
    public void TcPrOnly_is_untracked_right_apply_today()
    {
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/>"));

        // Pin: the shell digest participates in cell ContentHash → the table pair is Modified,
        // but the right tcPr is copied verbatim with no tcPrChange (the #250-noted gap).
        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "tcPrChange"));
        Assert.Single(body.Descendants(W + "tc").Elements(W + "tcPr").Elements(W + "shd"));

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Single(BodyOf(rejected).Descendants(W + "tc").Elements(W + "tcPr").Elements(W + "shd")); // reject ≠ left
    }

    // ------------------------------------------------------------------ trailing sectPr-only (w:pgSz)

    [Fact]
    public void TrailingSectPrOnly_is_silently_dropped_today()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/></w:sectPr>");

        Assert.Empty(DocxDiff.GetRevisions(left, right, ModeledOnly));

        // Pin: the LEFT trailing sectPr is preserved verbatim — the right page-size change is DROPPED,
        // so accept ≠ right at the sectPr level (the one silent-drop row of the family).
        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "sectPrChange"));
        var pgSz = body.Elements(W + "sectPr").Single().Element(W + "pgSz")!;
        Assert.Equal("12240", (string?)pgSz.Attribute(W + "w"));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var acceptedPgSz = BodyOf(accepted).Elements(W + "sectPr").Single().Element(W + "pgSz")!;
        Assert.Equal("12240", (string?)acceptedPgSz.Attribute(W + "w"));                    // accept ≠ right
    }

    // ------------------------------------------------------------------ the WmlComparer oracle

    [Fact]
    public void Oracle_WmlComparer_ignores_a_pPr_only_change()
    {
        // Rationale pin for the differential harness: the blessed oracle reports NOTHING for a
        // pPr-only change (it emits no pPrChange anywhere), so IR-side paragraph-scope format
        // revisions bucket as "IR more correct" / oracle-cannot-produce.
        var settings = new WmlComparerSettings();
        var compared = WmlComparer.Compare(PPrLeft, PPrRight, settings);
        var revisions = WmlComparer.GetRevisions(compared, settings);
        Assert.Empty(revisions);
        Assert.Empty(BodyOf(compared).Descendants(W + "pPrChange"));
    }
}
