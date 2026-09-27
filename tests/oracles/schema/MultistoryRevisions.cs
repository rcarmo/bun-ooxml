using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

// Development-only independent producer/readback. Expected documents are authored
// directly from the desired final state, never obtained by resolving source XML.
internal static class MultistoryRevisions
{
    private const string Ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly DateTime Stamp = new(2026, 9, 27, 0, 0, 0, DateTimeKind.Utc);
    internal static string[] Generate(string directory)
    {
        Directory.CreateDirectory(directory);
        var paths = new List<string>();
        foreach (var mode in new[] { "source", "accept", "reject", "unsupported" })
        {
            var path = Path.Combine(directory, $"multistory-{mode}.docx");
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                doc.ChangeIdOfPart(main, "rIdDocument");
                var header = main.AddNewPart<HeaderPart>("rIdHeader");
                var footer = main.AddNewPart<FooterPart>("rIdFooter");
                var notes = main.AddNewPart<FootnotesPart>("rIdFootnotes");
                var body = new W.Body(Paragraphs("Body", mode));
                body.Append(new W.Paragraph(new W.Run(new W.FootnoteReference { Id = 1 })));
                body.Append(new W.SectionProperties(
                    new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = "rIdHeader" },
                    new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = "rIdFooter" }));
                main.Document = new W.Document(body);
                header.Header = new W.Header(Paragraphs("Header", mode));
                footer.Footer = new W.Footer(Paragraphs("Footer", mode));
                var note = new W.Footnote(Paragraphs("Footnote", mode)) { Id = 1 };
                if (mode == "unsupported")
                    note.Append(new W.Paragraph(new W.ParagraphProperties(
                        new W.ParagraphPropertiesChange(new W.PreviousParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }))
                        { Id = "900", Author = "SDK author", Date = Stamp }), Run("Unsupported paragraph history")));
                notes.Footnotes = new W.Footnotes(
                    new W.Footnote(new W.Paragraph(new W.Run(new W.SeparatorMark()))) { Id = -1, Type = W.FootnoteEndnoteValues.Separator },
                    new W.Footnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) { Id = 0, Type = W.FootnoteEndnoteValues.ContinuationSeparator }, note);
                foreach (var root in new OpenXmlPartRootElement[] { main.Document, header.Header, footer.Footer, notes.Footnotes })
                {
                    root.AddNamespaceDeclaration("w", Ns);
                    root.Save();
                }
            }
            paths.Add(path);
        }
        return paths.ToArray();
    }

    private static W.Run Run(string text, bool bold = false, bool italic = false)
    {
        var run = new W.Run();
        if (bold || italic) { var pr = new W.RunProperties(); if (bold) pr.Append(new W.Bold()); if (italic) pr.Append(new W.Italic()); run.Append(pr); }
        run.Append(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }
    private static OpenXmlElement[] Paragraphs(string label, string mode)
    {
        var revision = mode == "source" || mode == "unsupported";
        var paragraphs = new List<OpenXmlElement> { new W.Paragraph(Run(label + " unchanged")) };
        var text = new W.Paragraph();
        if (revision)
        {
            text.Append(new W.DeletedRun(new W.Run(new W.DeletedText(label + " original") { Space = SpaceProcessingModeValues.Preserve }))
                { Id = "1", Author = "SDK author", Date = Stamp });
            text.Append(new W.InsertedRun(Run(label + " revised")) { Id = "2", Author = "SDK author", Date = Stamp });
        }
        else text.Append(Run(label + (mode == "accept" ? " revised" : " original")));
        paragraphs.Add(text);
        var formatRun = Run(label + " formatted", bold: mode != "reject", italic: mode == "reject");
        if (revision) formatRun.RunProperties!.Append(new W.RunPropertiesChange(new W.PreviousRunProperties(new W.Italic()))
            { Id = "3", Author = "SDK author", Date = Stamp });
        paragraphs.Add(new W.Paragraph(formatRun));
        var move = new W.Paragraph();
        if (revision)
        {
            move.Append(new W.MoveFromRangeStart { Id = "10", Name = "sdkMove", Author = "SDK author", Date = Stamp });
            move.Append(new W.MoveFromRun(Run(label + " moved", bold: true)) { Id = "11", Author = "SDK author", Date = Stamp });
            move.Append(new W.MoveFromRangeEnd { Id = "10" });
        }
        else if (mode == "reject") move.Append(Run(label + " moved", bold: true));
        move.Append(Run(label + " middle"));
        if (revision)
        {
            move.Append(new W.MoveToRangeStart { Id = "12", Name = "sdkMove", Author = "SDK author", Date = Stamp });
            move.Append(new W.MoveToRun(Run(label + " moved", bold: true)) { Id = "13", Author = "SDK author", Date = Stamp });
            move.Append(new W.MoveToRangeEnd { Id = "12" });
        }
        else if (mode == "accept") move.Append(Run(label + " moved", bold: true));
        paragraphs.Add(move);
        return paragraphs.ToArray();
    }

    internal static object Read(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart ?? throw new InvalidDataException("Missing main part");
        var parts = new List<(string kind, OpenXmlPart part, OpenXmlPartRootElement root)>
        { ("body", main, main.Document ?? throw new InvalidDataException("Missing document root")) };
        parts.AddRange(main.HeaderParts.Select(p => ("header", (OpenXmlPart)p, (OpenXmlPartRootElement)p.Header!)));
        parts.AddRange(main.FooterParts.Select(p => ("footer", (OpenXmlPart)p, (OpenXmlPartRootElement)p.Footer!)));
        if (main.FootnotesPart is { } notes) parts.Add(("footnotes", notes, notes.Footnotes!));
        return parts.Select(entry => new {
            kind = entry.kind, part = entry.part.Uri.ToString().TrimStart('/'),
            revisions = entry.root.Descendants().Where(e => e.NamespaceUri == Ns &&
                new[] { "ins", "del", "rPrChange", "pPrChange", "moveFrom", "moveTo" }.Contains(e.LocalName))
                .Select(e => new { kind = e.LocalName, id = e.GetAttribute("id", Ns).Value }).ToArray(),
            paragraphs = entry.root.Descendants<W.Paragraph>().Select(p => new {
                runs = p.Descendants<W.Run>().Select(r => new {
                    text = string.Concat(r.Descendants().Where(e => e is W.Text || e is W.DeletedText).Select(e => e.InnerText)),
                    bold = r.RunProperties?.Bold is { } b && (b.Val?.Value ?? true),
                    italic = r.RunProperties?.Italic is { } i && (i.Val?.Value ?? true),
                    elements = r.ChildElements.Select(e => e.LocalName).ToArray()
                }).ToArray()
            }).ToArray()
        }).ToArray();
    }
}
