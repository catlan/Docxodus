#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Small DOCX fixtures + readback helpers for the composite-merger tests. <see cref="Para"/>
/// builds an IrReader-clean one-section document (one single-run paragraph per supplied string),
/// delegating to <see cref="IrTestDocuments.Create"/> so the required StyleDefinitionsPart /
/// DocumentSettingsPart are present. <see cref="PlainText"/> and <see cref="MainPartXml"/> read a
/// document back for assertions used by later composite-merge tasks.
/// </summary>
internal static class Docs
{
    /// <summary>A one-section DOCX whose body holds one single-run paragraph per supplied string.</summary>
    public static WmlDocument Para(params string[] paragraphs) => IrTestDocuments.Create(paragraphs);

    /// <summary>Body paragraph text, paragraphs joined by newline (run text concatenated per paragraph).</summary>
    public static string PlainText(WmlDocument d)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var doc = XDocument.Parse(MainPartXml(d));
        var body = doc.Root?.Element(ns + "body");
        if (body is null)
            return string.Empty;
        var paras = body.Elements(ns + "p")
            .Select(p => string.Concat(p.Descendants(ns + "t").Select(t => t.Value)));
        return string.Join("\n", paras);
    }

    /// <summary>
    /// A normalized, table-aware structural projection of the MAIN document body, in document order: one
    /// tag per block — <c>"P:"+paragraph text</c> for each body <c>w:p</c>, and for each body <c>w:tbl</c>
    /// a <c>"TBL"</c> marker followed by a <c>"TC:"+cell text</c> tag per cell (descending row → cell →
    /// the cell's paragraph text). Unlike <see cref="PlainText"/> (which reads only the body's direct-child
    /// <c>w:p</c> and silently skips tables), this captures table presence and per-cell content, so a
    /// consolidate that corrupts or drops a table on the reject path produces a different projection.
    /// Walks only direct children of <c>w:body</c> so nested tables surface through their parent cell's text.
    /// Footnote text is intentionally NOT included here (footnote round-trip is asserted separately).
    /// </summary>
    public static string StructuralBody(WmlDocument d)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var doc = XDocument.Parse(MainPartXml(d));
        var body = doc.Root?.Element(ns + "body");
        if (body is null)
            return string.Empty;

        var sb = new StringBuilder();
        foreach (var block in body.Elements())
        {
            if (block.Name == ns + "p")
            {
                sb.Append("P:")
                  .Append(string.Concat(block.Descendants(ns + "t").Select(t => t.Value)))
                  .Append('\n');
            }
            else if (block.Name == ns + "tbl")
            {
                sb.Append("TBL\n");
                foreach (var tr in block.Elements(ns + "tr"))
                    foreach (var tc in tr.Elements(ns + "tc"))
                        sb.Append("TC:")
                          .Append(string.Concat(tc.Descendants(ns + "t").Select(t => t.Value)))
                          .Append('\n');
            }
            // Other body-level blocks (e.g. sectPr) are ignored — they are not part of the
            // content structure this oracle compares.
        }
        return sb.ToString();
    }

    /// <summary>The main document part XML as a string.</summary>
    public static string MainPartXml(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var main = wDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");
        using var partStream = main.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(partStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
