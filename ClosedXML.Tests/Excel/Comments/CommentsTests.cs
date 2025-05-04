using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Comments
{
    public class CommentsTests
    {
        [Test]
        public void CanConvertVmlPaletteEntriesToColors()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentsWithColorNamesAndIndexes.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var c = ws.FirstCellUsed();

            // None indicates an absence of a color
            var lineColor = c.GetComment().Style.ColorsAndLines.LineColor;
            Assert.Multiple(() =>
            {
                Assert.That(lineColor.ColorType, Is.EqualTo(XLColorType.Color));
                Assert.That(lineColor.Color.ToHex(), Is.EqualTo("00000000"));
            });

            var bgColor = c.GetComment().Style.ColorsAndLines.FillColor;
            Assert.Multiple(() =>
            {
                Assert.That(bgColor.ColorType, Is.EqualTo(XLColorType.Color));
                Assert.That(bgColor.Color.ToHex(), Is.EqualTo("FFFFFFE1"));
            });
        }

        [Test]
        public void CopyCommentStyle()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            string strExcelComment = "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "2) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "3) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "4) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "5) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "6) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "7) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "8) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;
            strExcelComment = strExcelComment + "9) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + Environment.NewLine;

            var cell = ws.Cell(2, 2).SetValue("Comment 1");

            cell.GetComment()
                .SetVisible(false)
                .AddText(strExcelComment);

            cell.GetComment()
                .Style
                .Alignment
                .SetAutomaticSize();

            cell.GetComment()
                .Style
                .ColorsAndLines
                .SetFillColor(XLColor.Red);

            ws.Row(1).InsertRowsAbove(1);

            Action<IXLCell> validate = c =>
            {
                Assert.Multiple(() =>
                {
                    Assert.That(c.GetComment().Style.Alignment.AutomaticSize, Is.True);
                    Assert.That(c.GetComment().Style.ColorsAndLines.FillColor, Is.EqualTo(XLColor.Red));
                });
            };

            validate(ws.Cell("B3"));

            ws.Column(1).InsertColumnsBefore(2);

            validate(ws.Cell("D3"));

            ws.Column(1).Delete();

            validate(ws.Cell("C3"));

            ws.Row(1).Delete();

            validate(ws.Cell("C2"));
        }

        [Test]
        public void EnsureUnaffectedCommentAndVmlPartIdsAndUris()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx"));
            using var ms = new MemoryStream();
            string commentPartId;
            string commentPartUri;

            string vmlPartId;
            string vmlPartUri;

            using (var ssd = SpreadsheetDocument.Open(stream, isEditable: false))
            {
                var wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
                var wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

                var wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
                commentPartId = wsp.GetIdOfPart(wscp);
                commentPartUri = wscp.Uri.ToString();

                var vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
                vmlPartId = wsp.GetIdOfPart(vmlp);
                vmlPartUri = vmlp.Uri.ToString();
            }

            stream.Position = 0;
            stream.CopyTo(ms);
            ms.Position = 0;

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.That(ws.FirstCell().HasComment, Is.True);

                wb.SaveAs(ms);
            }

            ms.Position = 0;

            using (var ssd = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
                var wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

                var wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
                Assert.Multiple(() =>
                {
                    Assert.That(wscp.Uri.ToString(), Is.EqualTo(commentPartUri));
                    Assert.That(wsp.GetIdOfPart(wscp), Is.EqualTo(commentPartId));
                });

                var vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
                Assert.Multiple(() =>
                {
                    Assert.That(vmlp.Uri.ToString(), Is.EqualTo(vmlPartUri));
                    Assert.That(wsp.GetIdOfPart(vmlp), Is.EqualTo(vmlPartId));
                });
            }
        }

        [Test]
        public void SavingDoesNotCauseTwoRootElements() // See #1157
        {
            using var ms = new MemoryStream();
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                wb.SaveAs(ms);
            }

            Assert.DoesNotThrow(() => new XLWorkbook(ms));
        }

        [Test]
        public void CanLoadCommentVisibility()
        {
            using var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Drawings\Comments\inputfile.xlsx"));
            using var workbook = new XLWorkbook(inputStream);
            var ws = workbook.Worksheets.First();

            Assert.That(ws.Cell("A1").GetComment().Visible, Is.True);
            Assert.False(ws.Cell("A4").GetComment().Visible);
        }

        [Test]
        public void Margins_are_converted_to_physical_length()
        {
            // Technically, it's insets on a textbox. Each comment uses a different unit, but all
            // should have same final dimension at left and top margin (easily visible in the
            // sheet). Tested units: in, cm, mm, pt, pc, emu, px, em, ex. Pixels are converted
            // through supplied DPI.
            // The last comment in vmlDrawing1 also has invalid units and number. These are
            // converted to 0, so we don't crash on load (Excel also ignores invalid values).
            var commentCells = new[] { "A1", "A7", "A16", "A22", "A28" };
            TestHelper.LoadAndAssert((_, ws) =>
            {
                foreach (var commentCell in commentCells)
                {
                    var cell = ws.Cell(commentCell);
                    Assert.That(cell.HasComment, Is.True);
                    var margins = cell.GetComment().Style.Margins;

                    Assert.That(margins.Left, Is.EqualTo(0.5));
                    Assert.That(margins.Top, Is.EqualTo(0.75));

                    Assert.That(margins.Right, Is.EqualTo(0));
                    Assert.That(margins.Bottom, Is.EqualTo(0));
                }
            }, @"Other\Comments\InsetsUnitConversion.xlsx", new LoadOptions { Dpi = new Point(120, 120) });
        }
    }
}
