using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Comments
{
    public class CommentsTests
    {
        [Test]
        public void CanGetColorFromIndex81()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentsWithIndexedColor81.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCellUsed();

                var xlColor = c.GetComment().Style.ColorsAndLines.LineColor;
                Assert.AreEqual(XLColorType.Indexed, xlColor.ColorType);
                Assert.AreEqual(81, xlColor.Indexed);

                var color = xlColor.Color.ToHex();
                Assert.AreEqual("FF000000", color);
            }
        }

        [Test]
        public void AddingCommentDoesNotAffectCollections()
        {
            var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
            ws.Cell("A1").SetValue(10);
            ws.Cell("A4").SetValue(10);
            ws.Cell("A5").SetValue(10);

            ws.Rows("1,4").Height = 20;

            Assert.AreEqual(2, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(3, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());

            ws.Cell("A4").GetComment().AddText("Comment");
            Assert.AreEqual(2, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(3, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());

            ws.Row(1).Delete();
            Assert.AreEqual(1, ws.Internals.RowsCollection.Count);
            Assert.AreEqual(2, ws.Internals.CellsCollection.RowsCollection.SelectMany(r => r.Value.Values).Count());
        }

        [Test]
        public void CopyCommentStyle()
        {
            using (var wb = new XLWorkbook())
            {
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
                    Assert.IsTrue(c.GetComment().Style.Alignment.AutomaticSize);
                    Assert.AreEqual(XLColor.Red, c.GetComment().Style.ColorsAndLines.FillColor);
                };

                validate(ws.Cell("B3"));

                ws.Column(1).InsertColumnsBefore(2);

                validate(ws.Cell("D3"));

                ws.Column(1).Delete();

                validate(ws.Cell("C3"));

                ws.Row(1).Delete();

                validate(ws.Cell("C2"));
            }
        }

        [Test]
        public void EnsureUnaffectedCommentAndVmlPartIdsAndUris()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")))
            using (var ms = new MemoryStream())
            {
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
                    Assert.IsTrue(ws.FirstCell().HasComment);

                    wb.SaveAs(ms);
                }

                ms.Position = 0;

                using (var ssd = SpreadsheetDocument.Open(ms, isEditable: false))
                {
                    var wbp = ssd.GetPartsOfType<WorkbookPart>().Single();
                    var wsp = wbp.GetPartsOfType<WorksheetPart>().Last();

                    var wscp = wsp.GetPartsOfType<WorksheetCommentsPart>().Single();
                    Assert.AreEqual(commentPartUri, wscp.Uri.ToString());
                    Assert.AreEqual(commentPartId, wsp.GetIdOfPart(wscp));

                    var vmlp = wsp.GetPartsOfType<VmlDrawingPart>().Single();
                    Assert.AreEqual(vmlPartUri, vmlp.Uri.ToString());
                    Assert.AreEqual(vmlPartId, wsp.GetIdOfPart(vmlp));
                }
            }
        }

        [Test]
        public void SavingDoesNotCauseTwoRootElements() // See #1157
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentAndButton.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    wb.SaveAs(ms);
                }

                Assert.DoesNotThrow(() => new XLWorkbook(ms));
            }
        }

        [Test]
        public void CanLoadCommentVisibility()
        {
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Drawings\Comments\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                var ws = workbook.Worksheets.First();

                Assert.True(ws.Cell("A1").GetComment().Visible);
                Assert.False(ws.Cell("A4").GetComment().Visible);
            }
        }
    }
}
