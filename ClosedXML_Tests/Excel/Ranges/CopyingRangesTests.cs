using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;
using Color = System.Drawing.Color;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class CopyingRangesTests
    {
        [Test]
        public void CopyingColumns()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            IXLColumn column1 = ws.Column(1);
            column1.Cell(1).Style.Fill.SetBackgroundColor(XLColor.Red);
            column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.FromArgb(1, 1, 1));
            column1.Cell(3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#CCCCCC"));
            column1.Cell(4).Style.Fill.SetBackgroundColor(XLColor.FromIndex(26));
            column1.Cell(5).Style.Fill.SetBackgroundColor(XLColor.FromColor(Color.MediumSeaGreen));
            column1.Cell(6).Style.Fill.SetBackgroundColor(XLColor.FromName("Blue"));
            column1.Cell(7).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent3));

            ws.Cell(1, 2).Value = column1;
            ws.Cell(1, 3).Value = column1.Column(1, 7);

            IXLColumn column2 = ws.Column(2);
            Assert.AreEqual(XLColor.Red, column2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromArgb(1, 1, 1), column2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromHtml("#CCCCCC"), column2.Cell(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromIndex(26), column2.Cell(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromColor(Color.MediumSeaGreen),
                column2.Cell(5).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromName("Blue"), column2.Cell(6).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromTheme(XLThemeColor.Accent3), column2.Cell(7).Style.Fill.BackgroundColor);

            IXLColumn column3 = ws.Column(3);
            Assert.AreEqual(XLColor.Red, column3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromArgb(1, 1, 1), column3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromHtml("#CCCCCC"), column3.Cell(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromIndex(26), column3.Cell(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromColor(Color.MediumSeaGreen),
                column3.Cell(5).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromName("Blue"), column3.Cell(6).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromTheme(XLThemeColor.Accent3), column3.Cell(7).Style.Fill.BackgroundColor);
        }

        [Test]
        public void CopyingRows()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            IXLRow row1 = ws.Row(1);
            FillRow(row1);

            ws.Cell(2, 1).Value = row1;
            ws.Cell(3, 1).Value = row1.Row(1, 7);

            IXLRow row2 = ws.Row(2);
            Assert.AreEqual(XLColor.Red, row2.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromArgb(1, 1, 1), row2.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromHtml("#CCCCCC"), row2.Cell(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromIndex(26), row2.Cell(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromColor(Color.MediumSeaGreen), row2.Cell(5).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromName("Blue"), row2.Cell(6).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromTheme(XLThemeColor.Accent3), row2.Cell(7).Style.Fill.BackgroundColor);

            IXLRow row3 = ws.Row(3);
            Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromArgb(1, 1, 1), row3.Cell(2).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromHtml("#CCCCCC"), row3.Cell(3).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromIndex(26), row3.Cell(4).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromColor(Color.MediumSeaGreen), row3.Cell(5).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromName("Blue"), row3.Cell(6).Style.Fill.BackgroundColor);
            Assert.AreEqual(XLColor.FromTheme(XLThemeColor.Accent3), row3.Cell(7).Style.Fill.BackgroundColor);

            Assert.AreEqual(3, ws.ConditionalFormats.Count());
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B1").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula));
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B2:B2").Values.Any(v => v.Value.Value == "G2" && v.Value.IsFormula));
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B3:B3").Values.Any(v => v.Value.Value == "G3" && v.Value.IsFormula));
        }

        [Test]
        public void CopyingConditionalFormats()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            FillRow(ws.Row(1));
            FillRow(ws.Row(2));
            FillRow(ws.Row(3));

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            ws.Cell(5, 2).Value = ws.Row(2).Row(1, 7);

            Assert.AreEqual(2, ws.ConditionalFormats.Count());
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B3").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula));
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "C5:C5").Values.Any(v => v.Value.Value == "H5" && v.Value.IsFormula));
        }

        [Test]
        public void CopyingConditionalFormatsDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var format = ws1.Range("A1:J2").AddConditionalFormat();

            var address = format.Ranges
                        .First()
                        .FirstCell()
                        .CellRight(4)
                        .Address
                        .ToStringRelative();

            format.WhenEquals("=" + address)
                .Fill
                .SetBackgroundColor(XLColor.Blue);

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.FirstCell().Value = ws1.Range("B1:B4");

            Assert.AreEqual(1, ws2.ConditionalFormats.Count());
            Assert.IsTrue(ws2.ConditionalFormats.All(x => x.Ranges.All(s => s.Worksheet == ws2)), "A conditional format was created for another worksheet.");
            Assert.IsTrue(ws2.ConditionalFormats
                .Single(x => x.Range.RangeAddress.ToStringRelative() == "A1:A2")
                .Values.Any(v => v.Value.Value == "E1" && v.Value.IsFormula), "The formula has not been transferred correctly.");

            Assert.AreEqual("Sheet1", ws1.ConditionalFormats.First().Ranges.First().Worksheet.Name);
            Assert.AreEqual("Sheet2", ws2.ConditionalFormats.First().Ranges.First().Worksheet.Name);
            Assert.AreEqual("A1:J2", ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
            Assert.AreEqual("A1:A2", ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString());
        }

        [Test]
        public void CopyConditionalFormatColorScaleInRange()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet");

            ws.Row(1).Cell(1).AddConditionalFormat()
                .ColorScale()
                .LowestValue(XLColor.Teal)
                .HighestValue(XLColor.Orange);

            ws.Cell(5, 2).Value = ws.Range(1, 1, 1, 5);

            Assert.AreEqual(2, ws.ConditionalFormats.Count()); //existing format + copied.
            Assert.IsTrue(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B5:B5").ConditionalFormatType == XLConditionalFormatType.ColorScale);
        }

        private static void FillRow(IXLRow row1)
        {
            row1.Cell(1).Style.Fill.SetBackgroundColor(XLColor.Red);
            row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.FromArgb(1, 1, 1));
            row1.Cell(3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#CCCCCC"));
            row1.Cell(4).Style.Fill.SetBackgroundColor(XLColor.FromIndex(26));
            row1.Cell(5).Style.Fill.SetBackgroundColor(XLColor.FromColor(Color.MediumSeaGreen));
            row1.Cell(6).Style.Fill.SetBackgroundColor(XLColor.FromName("Blue"));
            row1.Cell(7).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent3));

            row1.Cell(2).AddConditionalFormat().WhenEquals("=" + row1.FirstCell().CellRight(6).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
        }
    }
}
