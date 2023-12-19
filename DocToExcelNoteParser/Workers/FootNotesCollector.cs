using System;
using DocToExcelNoteParser.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToExcelNoteParser.Workers
{
	public class FootNotesCollector
	{
        public IEnumerable<FootNoteToken> GetFootNotes()
        {
            var docxFiles = Directory.GetFiles(".", "*.docx");

            var footNotes = new List<FootNoteToken>();

            foreach (var file in docxFiles)
            {
                using var wordDoc = WordprocessingDocument.Open(file, false);

                var footnotesPart = wordDoc.MainDocumentPart!.FootnotesPart;
                if (footnotesPart != null)
                {
                    var footnotes = footnotesPart.Footnotes.Elements<Footnote>().ToList();

                    foreach (var paragraph in wordDoc.MainDocumentPart.Document.Body!.Elements<Paragraph>())
                    {
                        var previousRun = paragraph.Elements<Run>().FirstOrDefault();
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            var footnoteReference = run.Descendants<FootnoteReference>().FirstOrDefault();
                            if (footnoteReference != null)
                            {
                                var footnote = footnotes.FirstOrDefault(f => f.Id == footnoteReference.Id);
                                if (footnote != null)
                                {
                                    var footNoteName = previousRun!.InnerText;
                                    if (footNoteName.StartsWith(','))
                                        footNoteName = footNoteName[2..];

                                    var footNoteToken = new FootNoteToken(footNoteName, footnote.InnerText);
                                    footNotes.Add(footNoteToken);
                                }
                            }
                            previousRun = run;
                        }
                    }
                }
            }

            return footNotes;
        }
    }
}

