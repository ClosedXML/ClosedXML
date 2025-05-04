using ClosedXML.Excel;
using NUnit.Framework;
using System.Drawing;
using System.Linq;

namespace ClosedXML.Tests
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

            ws.Cell(1, 2).CopyFrom(column1);
            ws.Cell(1, 3).CopyFrom(column1.Column(1, 7));

            IXLColumn column2 = ws.Column(2);
            Assert.Multiple(() =>
            {
                Assert.That(column2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(1, 1, 1)));
                Assert.That(column2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromHtml("#CCCCCC")));
                Assert.That(column2.Cell(4).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(26)));
                Assert.That(column2.Cell(5).Style.Fill.BackgroundColor,
                    Is.EqualTo(XLColor.FromColor(Color.MediumSeaGreen)));
                Assert.That(column2.Cell(6).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromName("Blue")));
                Assert.That(column2.Cell(7).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent3)));
            });

            IXLColumn column3 = ws.Column(3);
            Assert.Multiple(() =>
            {
                Assert.That(column3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(column3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(1, 1, 1)));
                Assert.That(column3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromHtml("#CCCCCC")));
                Assert.That(column3.Cell(4).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(26)));
                Assert.That(column3.Cell(5).Style.Fill.BackgroundColor,
                    Is.EqualTo(XLColor.FromColor(Color.MediumSeaGreen)));
                Assert.That(column3.Cell(6).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromName("Blue")));
                Assert.That(column3.Cell(7).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent3)));
            });
        }

        [Test]
        public void CopyingRows()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet");

            IXLRow row1 = ws.Row(1);
            FillRow(row1);

            ws.Cell(2, 1).CopyFrom(row1);
            ws.Cell(3, 1).CopyFrom(row1.Row(1, 7));

            IXLRow row2 = ws.Row(2);
            Assert.Multiple(() =>
            {
                Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(1, 1, 1)));
                Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromHtml("#CCCCCC")));
                Assert.That(row2.Cell(4).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(26)));
                Assert.That(row2.Cell(5).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromColor(Color.MediumSeaGreen)));
                Assert.That(row2.Cell(6).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromName("Blue")));
                Assert.That(row2.Cell(7).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent3)));
            });

            IXLRow row3 = ws.Row(3);
            Assert.Multiple(() =>
            {
                Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
                Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(1, 1, 1)));
                Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromHtml("#CCCCCC")));
                Assert.That(row3.Cell(4).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(26)));
                Assert.That(row3.Cell(5).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromColor(Color.MediumSeaGreen)));
                Assert.That(row3.Cell(6).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromName("Blue")));
                Assert.That(row3.Cell(7).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent3)));

                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(3));
                Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B1").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula), Is.True);
                Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B2:B2").Values.Any(v => v.Value.Value == "G2" && v.Value.IsFormula), Is.True);
                Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B3:B3").Values.Any(v => v.Value.Value == "G3" && v.Value.IsFormula), Is.True);
            });
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

            ws.Cell(5, 2).CopyFrom(ws.Row(2).Row(1, 7));

            Assert.Multiple(() =>
            {
                Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(2));
                Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B3").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula), Is.True);
                Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "C5:C5").Values.Any(v => v.Value.Value == "H5" && v.Value.IsFormula), Is.True);
            });
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

            ws2.FirstCell().CopyFrom(ws1.Range("B1:B4"));

            Assert.Multiple(() =>
            {
                Assert.That(ws2.ConditionalFormats.Count(), Is.EqualTo(1));
                Assert.That(ws2.ConditionalFormats.All(x => x.Ranges.All(s => s.Worksheet == ws2)), Is.True, "A conditional format was created for another worksheet.");
                Assert.That(ws2.ConditionalFormats
                    .Single(x => x.Range.RangeAddress.ToStringRelative() == "A1:A2")
                    .Values.Any(v => v.Value.Value == "E1" && v.Value.IsFormula), Is.True, "The formula has not been transferred correctly.");

                Assert.That(ws1.ConditionalFormats.First().Ranges.First().Worksheet.Name, Is.EqualTo("Sheet1"));
                Assert.That(ws2.ConditionalFormats.First().Ranges.First().Worksheet.Name, Is.EqualTo("Sheet2"));
                Assert.That(ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(), Is.EqualTo("A1:J2"));
            });
            Assert.That(ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(), Is.EqualTo("A1:A2"));
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
