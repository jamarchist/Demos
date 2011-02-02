using System;
using System.Data;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
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

        [Test]
        public void CanGeneratePrettyReport()
        {
            var data = GetTestData();

            var atlanticData = GetDivision("Atlantic", data);
            var centralData = GetDivision("Central", data);

            var doc = new Document();
            var section = doc.AddSection();

            var title = section.AddParagraph("Standings");
            title.Format.Font.Size = Unit.FromPoint(16);
            title.Format.Font.Underline = Underline.Single;
            title.Format.Alignment = ParagraphAlignment.Center;
            title.Format.SpaceAfter = Unit.FromPoint(12);

            var atlanticHeading = section.AddParagraph("Atlantic Division");
            atlanticHeading.Format.Font.Bold = true;
            atlanticHeading.Format.Borders.Bottom.Color = Colors.Black;
            atlanticHeading.Format.Borders.Bottom.Width = Unit.FromPoint(1);
            section.AddParagraph();
            
            var atlanticTable = AddDivisionTableTo(section, "Atlantic");
            section.AddParagraph();
            
            var centralHeading = section.AddParagraph("Cental Division");
            centralHeading.Format.Font.Bold = true;
            centralHeading.Format.Borders.Bottom.Color = Colors.Black;
            centralHeading.Format.Borders.Bottom.Width = Unit.FromPoint(1);
            section.AddParagraph();

            var centralTable = AddDivisionTableTo(section, "Central");
            
            PopulateDivisionTable(atlanticTable, atlanticData);
            PopulateDivisionTable(centralTable, centralData);

            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = doc;
            pdfRenderer.RenderDocument();

            pdfRenderer.Save(@"standings.pdf");
        }

        private static Table AddDivisionTableTo(Section section, string divisionName)
        {
            var divisionTable = section.AddTable();
            divisionTable.Comment = divisionName;

            var id = divisionTable.AddColumn(Unit.FromInch(0.5));
            var name = divisionTable.AddColumn(Unit.FromInch(2.5));
            var description = divisionTable.AddColumn(Unit.FromInch(3));
            var value = divisionTable.AddColumn(Unit.FromInch(1));

            var headerRow = divisionTable.AddRow();
            headerRow.HeadingFormat = true;
            headerRow.Shading.Color = Colors.LightGray;

            headerRow.Cells[0].AddParagraph("Id");
            headerRow.Cells[1].AddParagraph("City");
            headerRow.Cells[2].AddParagraph("Nickname");
            headerRow.Cells[3].AddParagraph("Wins");

            headerRow.Borders.Width = Unit.FromPoint(1);

            return divisionTable;
        }

        private static void PopulateDivisionTable(Table table, DataRow[] source)
        {
            foreach (var row in source)
            {
                var pdfRow = table.AddRow();
                pdfRow.Borders.Width = Unit.FromPoint(1);
                pdfRow.Borders.Bottom.Color = Colors.LightGray;
                pdfRow.Borders.Right.Color = Colors.LightGray;
                pdfRow.Borders.Left.Color = Colors.LightGray;

                pdfRow.Borders.Top.Color = Array.IndexOf(source, row) == 0 ? Colors.Black : Colors.LightGray;

                pdfRow.Cells[0].AddParagraph(row["Id"].ToString());
                pdfRow.Cells[1].AddParagraph(row["Name"].ToString());
                pdfRow.Cells[2].AddParagraph(row["Description"].ToString());
                pdfRow.Cells[3].AddParagraph(row["Value"].ToString());
            }
        }

        private static DataTable GetTestData()
        {
            var table = new DataTable();

            var recordId = new DataColumn("Id", typeof (int));
            var recordName = new DataColumn("Name", typeof (string));
            var recordDescription = new DataColumn("Description", typeof (string));
            var recordCategory = new DataColumn("Category", typeof (string));
            var recordValue = new DataColumn("Value", typeof (int));

            table.Columns.Add(recordId);
            table.Columns.Add(recordName);
            table.Columns.Add(recordDescription);
            table.Columns.Add(recordCategory);
            table.Columns.Add(recordValue);

            table.Rows.Add(1, "Boston", "Celtics", "Atlantic", 34);
            table.Rows.Add(2, "New York", "Knicks", "Atlantic", 23);
            table.Rows.Add(3, "Philadelphia", "76ers", "Atlantic", 19);
            table.Rows.Add(4, "Toronto", "Raptors", "Atlantic", 13);
            table.Rows.Add(5, "New Jersey", "Nets", "Atlantic", 13);

            table.Rows.Add(6, "Chicago", "Bulls", "Central", 31);
            table.Rows.Add(7, "Indiana", "Pacers", "Central", 16);
            table.Rows.Add(8, "Milwaukee", "Bucks", "Central", 16);
            table.Rows.Add(9, "Detroit", "Pistons", "Central", 17);
            table.Rows.Add(10, "Cleveland", "Cavaliers", "Central", 8);

            return table;
        }

        private static DataRow[] GetDivision(string divisionName, DataTable source)
        {
            return (from DataRow row in source.Rows where row["Category"].ToString() == divisionName select row).ToArray();
        }
    }
}
