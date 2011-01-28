using System;
using System.Collections.Generic;
using System.Text;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using NUnit.Framework;

namespace PdfDocumentGenerator
{
    [TestFixture]
    public class PdfGenerationTests
    {
        [Test]
        public void CanGeneratePdfTable()
        {
            var doc = new Document();
            var section = doc.AddSection();
            var table = section.AddTable();

            var column1 = table.AddColumn();
            var column2 = table.AddColumn();

            var row1 = table.AddRow();
            row1.Cells[0].AddParagraph("Ryan");
            row1.Cells[1].AddParagraph("30");

            var row2 = table.AddRow();
            row2.Cells[0].AddParagraph("Joe");
            row2.Cells[1].AddParagraph("99");

            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = doc;
            pdfRenderer.RenderDocument();

            pdfRenderer.Save(@"pdfsharptest.pdf");
        }
    }
}
