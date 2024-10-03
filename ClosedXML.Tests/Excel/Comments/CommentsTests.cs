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
        public void CanConvertVmlPaletteEntriesToColors()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\CommentsWithColorNamesAndIndexes.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCellUsed();

                // None indicates an absence of a color
                var lineColor = c.GetComment().Style.ColorsAndLines.LineColor;
                Assert.AreEqual(XLColorType.Color, lineColor.ColorType);
                Assert.AreEqual("00000000", lineColor.Color.ToHex());

                var bgColor = c.GetComment().Style.ColorsAndLines.FillColor;
                Assert.AreEqual(XLColorType.Color, bgColor.ColorType);
                Assert.AreEqual("FFFFFFE1", bgColor.Color.ToHex());
            }
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

        [Test]
        [TestCase(ThreadedCommentLoading.Skip, "")]
        [TestCase(ThreadedCommentLoading.ConvertToNotes, @"This is a threaded commentThis is a reply.")]
        public void CanReadThreadedCommentNote (ThreadedCommentLoading threadedCommentLoading, string expectedComments)
        {
            Assert.Multiple (() =>
            {
                using var stream = TestHelper.GetStreamFromResource (TestHelper.GetResourcePath (@"TryToLoad\ThreadedComment.xlsx"));
                using var wb = new XLWorkbook (stream, new LoadOptions { ThreadedCommentLoading = threadedCommentLoading });
                AssertComment (expectedComments, wb);

                using var ms1 = new MemoryStream ();
                wb.SaveAs (ms1, true);


                using var wb2 = new XLWorkbook (ms1);
                AssertComment (expectedComments, wb2);

                using var ms2 = new MemoryStream ();
                wb.SaveAs (ms2, true);

                static void AssertComment (string expectedComments, XLWorkbook wb)
                {
                    var ws = wb.Worksheets.First ();
                    var c1 = ws.Cell ("A1");
                    Assert.AreEqual (expectedComments, c1.GetComment ().Text);
                    Assert.AreEqual ("tc={49C52447-16DF-491E-8BD1-273F700714C6}", c1.GetComment ().Author);

                    var c2 = ws.Cell ("A2");
                    Assert.AreEqual ("Author:\r\nA note", c2.GetComment ().Text);
                    Assert.AreEqual ("tc={49C52447-16DF-491E-8BD1-273F700714C6}", c1.GetComment ().Author);
                }
            });
        }

        [Test]
        public void CanRemoveCommentsWithoutAddingOthers_Regression ()
        {
            Assert.Multiple (() =>
            {
                using (var stream = new MemoryStream ())
                {
                    // arange
                    using (var wb = new XLWorkbook ())
                    {
                        var sheet = wb.AddWorksheet ("sheet1");

                        var a1 = sheet.Cell ("A1");
                        var b5 = sheet.Cell ("B5");

                        a1.SetValue ("test a1");
                        b5.SetValue ("test b5");

                        a1.GetComment().AddText ("no comment");

                        var cellsWithComments3 = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        Assert.That (cellsWithComments3.Length, Is.EqualTo (1));

                        wb.SaveAs (stream, true);
                    }

                    stream.Position = 0;

                    using (var wb = new XLWorkbook (stream))
                    {
                        var cellsWithComments = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        Assert.That (cellsWithComments.Length, Is.EqualTo (1));

                        cellsWithComments.ForEach (_ => _.Clear (XLClearOptions.Comments));

                        wb.Save ();
                    }

                    // assert
                    stream.Position = 0;

                    using (var wb = new XLWorkbook (stream))
                    {
                        var cellsWithComments = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        // BUG? when adding a1.SetValue ("test a1"); this will return at least one cell instead of none
                        Assert.That (cellsWithComments, Is.Empty);
                    }
                }
            });
        }
    }
}
