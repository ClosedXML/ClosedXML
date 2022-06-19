using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Comments
{
    public class CommentsTests
    {
        [Test]
        public void CanGetColorFromIndex81()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentsWithIndexedColor81.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var c = ws.FirstCellUsed();

            var xlColor = c.GetComment().Style.ColorsAndLines.LineColor;
            Assert.AreEqual(XLColorType.Indexed, xlColor.ColorType);
            Assert.AreEqual(81, xlColor.Indexed);

            var color = xlColor.Color.ToHex();
            Assert.AreEqual("FF000000", color);
        }

        [Test]
        public void AddingCommentDoesNotAffectCollections()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet() as XLWorksheet;
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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var strExcelComment = "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "1) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "2) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "3) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "4) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "5) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "6) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "7) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "8) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;
            strExcelComment = strExcelComment + "9) ABCDEFGHIJKLMNOPQRSTUVWXYZ ABC ABC ABC ABC ABC" + XLConstants.NewLine;

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

            static void validate(IXLCell c)
            {
                Assert.IsTrue(c.GetComment().Style.Alignment.AutomaticSize);
                Assert.AreEqual(XLColor.Red, c.GetComment().Style.ColorsAndLines.FillColor);
            }

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

            Assert.True(ws.Cell("A1").GetComment().Visible);
            Assert.False(ws.Cell("A4").GetComment().Visible);
        }

        [Test]
        public void CanRemoveCommentsWithoutAddingOthers() // see #1575
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.AddWorksheet("sheet1");

                //guard
                var cellsWithComments1 = wb.Worksheets.SelectMany(_ => _.CellsUsed(XLCellsUsedOptions.Comments)).ToArray();
                Assert.That(cellsWithComments1, Is.Empty);

                var a1 = sheet.Cell("A1");
                var b5 = sheet.Cell("B5");

                a1.SetValue("test a1");
                b5.SetValue("test b5");

                var cellsWithComments2 = wb.Worksheets.SelectMany(_ => _.CellsUsed(XLCellsUsedOptions.Comments)).ToArray();
                Assert.That(cellsWithComments2, Is.Empty);

                a1.GetComment().AddText("no comment");

                //guard
                var cellsWithComments3 = wb.Worksheets.SelectMany(_ => _.CellsUsed(XLCellsUsedOptions.Comments)).ToArray();
                Assert.That(cellsWithComments3.Length, Is.EqualTo(1));

                wb.SaveAs(ms, true);
            }

            ms.Position = 0;

            using (var wb = new XLWorkbook(ms))
            {
                var cellsWithComments = wb.Worksheets.SelectMany(_ => _.CellsUsed(XLCellsUsedOptions.Comments)).ToArray();

                Assert.That(cellsWithComments.Length, Is.EqualTo(1));

                // act
                cellsWithComments.ForEach(_ => _.Clear(XLClearOptions.Comments));

                wb.Save();
            }

            ms.Position = 0;

            using (var wb = new XLWorkbook(ms))
            {
                // assert
                var cellsWithComments = wb.Worksheets.SelectMany(_ => _.Cells(true, XLCellsUsedOptions.Comments)).ToArray();

                Assert.That(cellsWithComments, Is.Empty);
            }
        }

        [Test]
        public void CanDeleteFormattedNote()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentFormatted.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            var cellsWithComments = ws.CellsUsed(XLCellsUsedOptions.Comments).ToArray();

            Assert.That(cellsWithComments.Length, Is.EqualTo(2));

            Assert.That(cellsWithComments[0].GetComment().Text, Is.EqualTo(@"normal Note"));
            Assert.That(cellsWithComments[1].GetComment().Text, Is.EqualTo("Author:\r\nboldAndUnderlinenormal bold italic normal"));

            cellsWithComments[0].Clear(XLClearOptions.Comments);
            cellsWithComments[1].Clear(XLClearOptions.Comments);

            cellsWithComments = ws.CellsUsed(XLCellsUsedOptions.Comments).ToArray();
            // TODO: this breaks when it should not
            Assert.That(cellsWithComments.Length, Is.EqualTo(0));
        }

        [Test]
        public void CanReadThreadedCommentNote()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\ThreadedComment.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var c = ws.FirstCellUsed();

            Assert.AreEqual(c.GetComment().Text, @"[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    This is a threaded comment.
Reply:
    This is a reply.");
        }
    }
}