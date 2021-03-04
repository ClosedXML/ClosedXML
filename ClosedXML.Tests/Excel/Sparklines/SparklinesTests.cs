using ClosedXML.Examples.Sparklines;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Sparklines
{
    [TestFixture]
    public class SparklinesTests
    {
        #region Add sparklines

        [Test]
        public void CannotCreateSparklineGroupsWithoutWorksheet()
        {
            TestDelegate action = () => new XLSparklineGroups(null);
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void CannotCreateSparklineGroupWithoutWorksheet()
        {
            TestDelegate action = () => new XLSparklineGroup(null);
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void CannotCreateSparklineWithoutGroup()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            TestDelegate action = () => new XLSparkline(null, ws.Cell("A1"), ws.Range("A2:A5"));
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void CannotCreateSparklineWithoutLocation()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var group = new XLSparklineGroup(ws);
            TestDelegate action = () => new XLSparkline(group, null, ws.Range("A2:A5"));
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void CanCreateInvalidSparklineWithoutSourceData()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var group = new XLSparklineGroup(ws);
            var sparkline = new XLSparkline(group, ws.FirstCell(), null);
            Assert.IsFalse(sparkline.IsValid);
        }

        [Test]
        public void CanAddSparklineGroupForSingleCell()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add(new XLSparklineGroup(ws, "A1", "B1:E1"));
            ws.SparklineGroups.Add("A2", "B2:E2");
            ws.SparklineGroups.Add(ws.Cell("A3"), ws.Range("B3:E3"));

            Assert.AreEqual(3, ws.SparklineGroups.Count());

            Assert.AreEqual("A1", ws.SparklineGroups.ElementAt(0).Single().Location.Address.ToString());
            Assert.AreEqual("A2", ws.SparklineGroups.ElementAt(1).Single().Location.Address.ToString());
            Assert.AreEqual("A3", ws.SparklineGroups.ElementAt(2).Single().Location.Address.ToString());

            Assert.AreEqual("B1:E1", ws.SparklineGroups.ElementAt(0).Single().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:E2", ws.SparklineGroups.ElementAt(1).Single().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B3:E3", ws.SparklineGroups.ElementAt(2).Single().SourceData.RangeAddress.ToString());

            Assert.IsTrue(ws.SparklineGroups.All(g => g.Worksheet == ws));
        }

        [Test]
        public void CanAddSparklineGroupForVerticalRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add(ws.Range("A1:A3"), ws.Range("B1:E3"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());

            Assert.AreEqual("A1", ws.SparklineGroups.Single().ElementAt(0).Location.Address.ToString());
            Assert.AreEqual("A2", ws.SparklineGroups.Single().ElementAt(1).Location.Address.ToString());
            Assert.AreEqual("A3", ws.SparklineGroups.Single().ElementAt(2).Location.Address.ToString());

            Assert.AreEqual("B1:E1", ws.SparklineGroups.Single().ElementAt(0).SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:E2", ws.SparklineGroups.Single().ElementAt(1).SourceData.RangeAddress.ToString());
            Assert.AreEqual("B3:E3", ws.SparklineGroups.Single().ElementAt(2).SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CanAddSparklineGroupForHorizontalRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add(ws.Range("A1:C1"), ws.Range("A2:C4"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());

            Assert.AreEqual("A1", ws.SparklineGroups.Single().ElementAt(0).Location.Address.ToString());
            Assert.AreEqual("B1", ws.SparklineGroups.Single().ElementAt(1).Location.Address.ToString());
            Assert.AreEqual("C1", ws.SparklineGroups.Single().ElementAt(2).Location.Address.ToString());

            Assert.AreEqual("A2:A4", ws.SparklineGroups.Single().ElementAt(0).SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:B4", ws.SparklineGroups.Single().ElementAt(1).SourceData.RangeAddress.ToString());
            Assert.AreEqual("C2:C4", ws.SparklineGroups.Single().ElementAt(2).SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CannotAddSparklineForNonLinearRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:C2"), ws.Range("A3:C4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("locationRange must have either a single row or a single column", message);
        }

        [Test]
        public void CannotAddSparklineWhenRangesHaveDifferentWidths()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:C1"), ws.Range("A3:D4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("locationRange and sourceDataRange must have the same width", message);
        }

        [Test]
        public void CannotAddSparklineWhenRangesHaveDifferentHeights()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:A3"), ws.Range("B1:B4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("locationRange and sourceDataRange must have the same height", message);
        }

        [Test]
        public void CannotAddSparklineForCellWhenDataRangeIsNotLinear()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:A1"), ws.Range("B1:C4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("SourceData range must have either a single row or a single column", message);
        }

        [Test]
        public void CanAddSparklineToExistingGroup()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            var group = new XLSparklineGroup(ws);

            group.Add("A2", "B2:E2");
            group.Add(ws.Cell("A3"), ws.Range("B3:E3"));

            Assert.AreEqual(0, ws.SparklineGroups.Count());

            Assert.AreEqual("A2", group.ElementAt(0).Location.Address.ToString());
            Assert.AreEqual("A3", group.ElementAt(1).Location.Address.ToString());

            Assert.AreEqual("B2:E2", group.ElementAt(0).SourceData.RangeAddress.ToString());
            Assert.AreEqual("B3:E3", group.ElementAt(1).SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CannotAddSparklineGroupFromDifferentWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");

            var group = new XLSparklineGroup(ws1);

            TestDelegate action = () => ws2.SparklineGroups.Add(group);

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("The specified sparkline group belongs to the different worksheet", message);
        }

        [Test]
        public void CannotAddSparklineFromDifferentWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");

            var group = new XLSparklineGroup(ws1);

            TestDelegate action = () => group.Add(ws2.Cell("A3"), ws1.Range("B3:E3"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("The specified sparkline belongs to the different worksheet", message);
        }

        [Test]
        public void AddSparklineToSameCellOverwritesItWhenSameGroup()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            var group = ws.SparklineGroups.Add("A1", "B1:E1");
            group.Add("A1", "B2:E2");

            Assert.AreEqual(1, group.Count());

            Assert.AreEqual("A1", group.Single().Location.Address.ToString());
            Assert.AreEqual("B2:E2", group.Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void AddSparklineToSameCellOverwritesItWhenDifferentGroup()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1", "B1:E1");
            ws.SparklineGroups.Add("A1", "B2:E2");

            Assert.AreEqual(2, ws.SparklineGroups.Count());
            Assert.IsFalse(ws.SparklineGroups.First().Any());
            Assert.AreEqual("A1", ws.SparklineGroups.Last().Single().Location.Address.ToString());
            Assert.AreEqual("B2:E2", ws.SparklineGroups.Last().Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CanAddSparklineReferringToDifferentWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws3 = wb.AddWorksheet("Sheet 3");

            var group = ws1.SparklineGroups.Add("A1", "'Sheet 3'!B1:F1");

            Assert.AreSame(ws3, group.Single().SourceData.Worksheet);
        }

        #endregion Add sparklines

        #region Get sparklines

        [TestCase("A2", "B2:Z2")]
        [TestCase("A50", "B50:Z50")]
        [TestCase("A100", "B100:Z100")]
        [TestCase("B1", "B2:B100")]
        [TestCase("K1", "K2:K100")]
        [TestCase("Z1", "Z2:Z100")]
        public void CanGetSparklineForExistingCell(string cellAddress, string expectedSourceDataRange)
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A2:A100", "B2:Z100");
            ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

            var sp = ws.SparklineGroups.GetSparkline(ws.Cell(cellAddress));
            Assert.IsNotNull(sp);
            Assert.AreEqual(cellAddress, sp.Location.Address.ToString());
            Assert.AreEqual(expectedSourceDataRange, sp.SourceData.RangeAddress.ToString());
        }

        [TestCase("A1")]
        [TestCase("B2")]
        [TestCase("A101")]
        [TestCase("AA1")]
        public void CannotGetSparklineForNonExistingCell(string cellAddress)
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A2:A100", "B2:Z100");
            ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

            var sp = ws.SparklineGroups.GetSparkline(ws.Cell(cellAddress));
            Assert.IsNull(sp);
        }

        [Test]
        public void CanGetSparklinesForRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A2:A100", "B2:Z100");
            ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

            var sparklines1 = ws.SparklineGroups.GetSparklines(ws.Range("A1:B2"));
            var sparklines2 = ws.SparklineGroups.GetSparklines(ws.Range("B2:E4"));
            var sparklines3 = ws.SparklineGroups.GetSparklines(ws.Range("A1:Z100"));
            var sparklines4 = ws.SparklineGroups.GetSparklines(ws.Range("A:A"));
            var sparklines5 = ws.SparklineGroups.GetSparklines(ws.Range("1:1"));

            Assert.AreEqual(2, sparklines1.Count());
            Assert.AreEqual(0, sparklines2.Count());
            Assert.AreEqual(99 + 25, sparklines3.Count());
            Assert.AreEqual(99, sparklines4.Count());
            Assert.AreEqual(25, sparklines5.Count());

            Assert.AreEqual("A2", sparklines1.First().Location.Address.ToString());
            Assert.AreEqual("B1", sparklines1.Last().Location.Address.ToString());
            Assert.AreEqual("B2:Z2", sparklines1.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:B100", sparklines1.Last().SourceData.RangeAddress.ToString());
        }

        #endregion Get sparklines

        #region Remove sparklines

        [Test]
        public void CanRemoveSparklineFromCell()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A3", "B1:Z3");
            ws.SparklineGroups.Remove(ws.Cell("A2"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(2, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("A3", ws.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B3:Z3", ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CanRemoveSparklineFromRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A5", "B1:Z5");
            ws.SparklineGroups.Remove(ws.Range("A2:D4"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(2, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("A5", ws.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B5:Z5", ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void RemoveSparklineFromEmptyCellDoesNothing()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws.SparklineGroups.Remove(ws.Cell("F2"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(2, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("A2", ws.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:Z2", ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString());
        }

        #endregion Remove sparklines

        #region Change sparklines

        [Test]
        public void CanChangeSparklineLocationInsideWorksheet()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws.SparklineGroups.Single().Last().SetLocation(ws.Cell("F2"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(2, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("F2", ws.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B2:Z2", ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString());
            Assert.IsTrue(ws.Cell("A1").HasSparkline);
            Assert.IsFalse(ws.Cell("A2").HasSparkline);
            Assert.IsTrue(ws.Cell("F2").HasSparkline);
        }

        [Test]
        public void ChangeSparklineLocationOverwritesExistingSparklineSameGroup()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws.SparklineGroups.Single().Last().SetLocation(ws.Cell("A1"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(1, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().Single().Location.Address.ToString());
            Assert.AreEqual("B2:Z2", ws.SparklineGroups.Single().Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void ChangeSparklineLocationOverwritesExistingSparklineDifferentGroups()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws.SparklineGroups.Add("A3", "B3:Z3");
            ws.SparklineGroups.Last().Single().SetLocation(ws.Cell("A2"));

            Assert.AreEqual(2, ws.SparklineGroups.Count());
            Assert.AreEqual(1, ws.SparklineGroups.First().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.First().Single().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.First().Single().SourceData.RangeAddress.ToString());
            Assert.AreEqual(1, ws.SparklineGroups.Last().Count());
            Assert.AreEqual("A2", ws.SparklineGroups.Last().Single().Location.Address.ToString());
            Assert.AreEqual("B3:Z3", ws.SparklineGroups.Last().Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CannotChangeSparklineLocationToAnotherWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");

            var group = ws1.SparklineGroups.Add("A1:A2", "B1:Z2");

            TestDelegate action = () => group.First().SetLocation(ws2.FirstCell());

            var message = Assert.Throws<InvalidOperationException>(action).Message;
            Assert.AreEqual("Cannot move the sparkline to a different worksheet", message);
        }

        [Test]
        public void CanChangeSparklineSourceDataInsideWorksheet()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws.SparklineGroups.Single().Last().SetSourceData(ws.Range("D4:D50"));

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(2, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("A2", ws.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("B1:Z1", ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("D4:D50", ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CannotChangeSparklineSourceDataToNonLinearRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1", "B1:Z1");
            var sparkline = group.Single();

            TestDelegate action = () => sparkline.SetSourceData(ws.Range("B1:Z2"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("SourceData range must have either a single row or a single column", message);
        }

        [Test]
        public void CanChangeSparklineStyle()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1", "B1:Z1");

            group.Style = XLSparklineTheme.Colorful1;

            Assert.AreEqual(XLColor.FromHtml("FF5F5F5F"), group.Style.SeriesColor);
            Assert.AreEqual(XLColor.FromHtml("FFFFB620"), group.Style.NegativeColor);
            Assert.AreEqual(XLColor.FromHtml("FFD70077"), group.Style.MarkersColor);
            Assert.AreEqual(XLColor.FromHtml("FF56BE79"), group.Style.HighMarkerColor);
            Assert.AreEqual(XLColor.FromHtml("FFFF5055"), group.Style.LowMarkerColor);
            Assert.AreEqual(XLColor.FromHtml("FF5687C2"), group.Style.FirstMarkerColor);
            Assert.AreEqual(XLColor.FromHtml("FF359CEB"), group.Style.LastMarkerColor);
        }

        [Test]
        public void ChangeSparklineStyleDoesNotAffectOriginal()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1", "B1:Z1");
            group.Style = XLSparklineTheme.Colorful1;

            group.Style.NegativeColor = XLColor.Red;

            Assert.AreEqual(XLColor.Red, group.Style.NegativeColor);
            Assert.AreNotEqual(XLColor.Red, XLSparklineTheme.Colorful1.NegativeColor);
        }

        [Test]
        public void CannotSetSparklineStyleToNull()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1", "B1:Z1");

            TestDelegate action = () => group.Style = null;

            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void SparklinesShiftOnRowInsert()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group1 = ws.SparklineGroups.Add("B2", "D4:F4");
            var group2 = ws.SparklineGroups.Add("B3", "D4:D8");
            var group3 = ws.SparklineGroups.Add("B4", "E1:E8");

            ws.Row(2).InsertRowsBelow(3);

            Assert.AreEqual("B2", group1.First().Location.Address.ToString());
            Assert.AreEqual("D7:F7", group1.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B6", group2.First().Location.Address.ToString());
            Assert.AreEqual("D7:D11", group2.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B7", group3.First().Location.Address.ToString());
            Assert.AreEqual("E1:E11", group3.First().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklinesShiftOnRowDelete()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group1 = ws.SparklineGroups.Add("B2", "D7:F7");
            var group2 = ws.SparklineGroups.Add("B6", "D7:D11");
            var group3 = ws.SparklineGroups.Add("B7", "E1:E11");

            ws.Rows(3, 5).Delete();

            Assert.AreEqual("B2", group1.First().Location.Address.ToString());
            Assert.AreEqual("D4:F4", group1.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B3", group2.First().Location.Address.ToString());
            Assert.AreEqual("D4:D8", group2.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B4", group3.First().Location.Address.ToString());
            Assert.AreEqual("E1:E8", group3.First().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklinesShiftOnColumnInsert()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group1 = ws.SparklineGroups.Add("B2", "D4:F4");
            var group2 = ws.SparklineGroups.Add("C3", "D4:D8");
            var group3 = ws.SparklineGroups.Add("D4", "A4:E4");

            ws.Column(2).InsertColumnsAfter(3);

            Assert.AreEqual("B2", group1.First().Location.Address.ToString());
            Assert.AreEqual("G4:I4", group1.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("F3", group2.First().Location.Address.ToString());
            Assert.AreEqual("G4:G8", group2.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("G4", group3.First().Location.Address.ToString());
            Assert.AreEqual("A4:H4", group3.First().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklinesShiftOnColumnDelete()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group1 = ws.SparklineGroups.Add("B2", "G4:I4");
            var group2 = ws.SparklineGroups.Add("F3", "G4:G8");
            var group3 = ws.SparklineGroups.Add("G4", "A4:H4");

            ws.Columns(3, 5).Delete();

            Assert.AreEqual("B2", group1.First().Location.Address.ToString());
            Assert.AreEqual("D4:F4", group1.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("C3", group2.First().Location.Address.ToString());
            Assert.AreEqual("D4:D8", group2.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("D4", group3.First().Location.Address.ToString());
            Assert.AreEqual("A4:E4", group3.First().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklineRemovedWhenColumnDeleted()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:B1", "C2:D6");

            ws.Column(2).Delete();

            Assert.AreEqual(1, group.Count());
            Assert.AreEqual("A1", group.Single().Location.Address.ToString());
            Assert.AreEqual("B2:B6", group.Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklineRemovedWhenRowDeleted()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:A2", "C3:F4");

            ws.Row(2).Delete();

            Assert.AreEqual(1, group.Count());
            Assert.AreEqual("A1", group.Single().Location.Address.ToString());
            Assert.AreEqual("C2:F2", group.Single().SourceData.RangeAddress.ToString());
        }

        [Test]
        public void SparklineRemovedWhenShiftedTooFarRight()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("XFD1", "A1:Z1");

            ws.Column(1).InsertColumnsBefore(1);

            Assert.AreEqual(0, group.Count());
        }

        [Test]
        public void SparklineRemovedWhenShiftedTooFarDown()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1048576", "A1:Z1");

            ws.Row(1).InsertRowsAbove(1);

            Assert.AreEqual(0, group.Count());
        }

        [Test]
        public void SparklineRangeInvalidatedWhenDeleted()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:B1", "C2:D6");

            ws.Column(4).Delete();

            Assert.AreEqual(2, group.Count());
            Assert.AreEqual("A1", group.First().Location.Address.ToString());
            Assert.AreEqual("C2:C6", group.First().SourceData.RangeAddress.ToString());
            Assert.AreEqual("B1", group.Last().Location.Address.ToString());
            Assert.IsFalse(group.Last().SourceData.RangeAddress.IsValid);
        }

        #endregion Change sparklines

        #region Load and save sparkline groups

        [Test]
        public void CanChangeSaveAndLoadSparklineGroup()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet 1");
                    var originalGroup = ws.SparklineGroups.Add("A1:A3", "B1:Z3")
                        .SetDateRange(ws.Range("B4:Z4"))
                        .SetLineWeight(5.5)
                        .SetDisplayHidden(true)
                        .SetShowMarkers(XLSparklineMarkers.FirstPoint | XLSparklineMarkers.LastPoint |
                                        XLSparklineMarkers.HighPoint | XLSparklineMarkers.LowPoint |
                                        XLSparklineMarkers.NegativePoints | XLSparklineMarkers.Markers)
                        .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
                        .SetType(XLSparklineType.Stacked);

                    originalGroup.HorizontalAxis
                        .SetColor(XLColor.AirForceBlue)
                        .SetVisible(true)
                        .SetRightToLeft(true);

                    originalGroup.VerticalAxis
                        .SetManualMax(6.6)
                        .SetManualMin(1.2)
                        .SetMaxAxisType(XLSparklineAxisMinMax.Custom)
                        .SetMinAxisType(XLSparklineAxisMinMax.Custom);

                    originalGroup.Style
                        .SetFirstMarkerColor(XLColor.AliceBlue)
                        .SetHighMarkerColor(XLColor.Alizarin)
                        .SetLastMarkerColor(XLColor.Almond)
                        .SetLowMarkerColor(XLColor.Amaranth)
                        .SetMarkersColor(XLColor.Amber)
                        .SetNegativeColor(XLColor.AmberSaeEce)
                        .SetSeriesColor(XLColor.AmericanRose);

                    AssertGroupIsValid(originalGroup);
                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();

                    Assert.AreEqual(1, ws.SparklineGroups.Count());
                    AssertGroupIsValid(ws.SparklineGroups.Single());
                }
            }

            void AssertGroupIsValid(IXLSparklineGroup group)
            {
                Assert.AreEqual(3, group.Count());

                Assert.AreEqual("A1", group.ElementAt(0).Location.Address.ToString());
                Assert.AreEqual("A2", group.ElementAt(1).Location.Address.ToString());
                Assert.AreEqual("A3", group.ElementAt(2).Location.Address.ToString());

                Assert.AreEqual("B1:Z1", group.ElementAt(0).SourceData.RangeAddress.ToString());
                Assert.AreEqual("B2:Z2", group.ElementAt(1).SourceData.RangeAddress.ToString());
                Assert.AreEqual("B3:Z3", group.ElementAt(2).SourceData.RangeAddress.ToString());

                Assert.AreEqual("B4:Z4", group.DateRange.RangeAddress.ToString());

                Assert.AreEqual(XLColor.AliceBlue, group.Style.FirstMarkerColor);
                Assert.AreEqual(XLColor.Alizarin, group.Style.HighMarkerColor);
                Assert.AreEqual(XLColor.Almond, group.Style.LastMarkerColor);
                Assert.AreEqual(XLColor.Amaranth, group.Style.LowMarkerColor);
                Assert.AreEqual(XLColor.Amber, group.Style.MarkersColor);
                Assert.AreEqual(XLColor.AmberSaeEce, group.Style.NegativeColor);
                Assert.AreEqual(XLColor.AmericanRose, group.Style.SeriesColor);
                Assert.IsTrue(group.DisplayHidden);
                Assert.AreEqual(5.5, group.LineWeight, XLHelper.Epsilon);
                Assert.AreEqual(XLDisplayBlanksAsValues.Zero, group.DisplayEmptyCellsAs);
                Assert.AreEqual(XLSparklineType.Stacked, group.Type);

                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.FirstPoint));
                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.LastPoint));
                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.HighPoint));
                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.LowPoint));
                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.NegativePoints));
                Assert.IsTrue(group.ShowMarkers.HasFlag(XLSparklineMarkers.Markers));

                Assert.AreEqual(XLColor.AirForceBlue, group.HorizontalAxis.Color);
                Assert.IsTrue(group.HorizontalAxis.IsVisible);
                Assert.IsTrue(group.HorizontalAxis.RightToLeft);
                Assert.IsTrue(group.HorizontalAxis.DateAxis);

                Assert.AreEqual(6.6, group.VerticalAxis.ManualMax, XLHelper.Epsilon);
                Assert.AreEqual(1.2, group.VerticalAxis.ManualMin, XLHelper.Epsilon);
                Assert.AreEqual(XLSparklineAxisMinMax.Custom, group.VerticalAxis.MaxAxisType);
                Assert.AreEqual(XLSparklineAxisMinMax.Custom, group.VerticalAxis.MinAxisType);
            }
        }

        [Test]
        public void CanLoadSparklines()
        {
            using (var ms = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx")))
            using (var wb = new XLWorkbook(ms))
            {
                Assert.IsTrue(wb.Worksheets.All(ws => ws.SparklineGroups.Count() == 6));
            }
        }

        [TestCase("Accent!B1", nameof(XLSparklineTheme.Accent1))]
        [TestCase("Accent!B2", nameof(XLSparklineTheme.Accent2))]
        [TestCase("Accent!B3", nameof(XLSparklineTheme.Accent3))]
        [TestCase("Accent!B4", nameof(XLSparklineTheme.Accent4))]
        [TestCase("Accent!B5", nameof(XLSparklineTheme.Accent5))]
        [TestCase("Accent!B6", nameof(XLSparklineTheme.Accent6))]
        [TestCase("'Accent Darker 25%'!B1", nameof(XLSparklineTheme.Accent1Darker25))]
        [TestCase("'Accent Darker 25%'!B2", nameof(XLSparklineTheme.Accent2Darker25))]
        [TestCase("'Accent Darker 25%'!B3", nameof(XLSparklineTheme.Accent3Darker25))]
        [TestCase("'Accent Darker 25%'!B4", nameof(XLSparklineTheme.Accent4Darker25))]
        [TestCase("'Accent Darker 25%'!B5", nameof(XLSparklineTheme.Accent5Darker25))]
        [TestCase("'Accent Darker 25%'!B6", nameof(XLSparklineTheme.Accent6Darker25))]
        [TestCase("'Accent Darker 50%'!B1", nameof(XLSparklineTheme.Accent1Darker50))]
        [TestCase("'Accent Darker 50%'!B2", nameof(XLSparklineTheme.Accent2Darker50))]
        [TestCase("'Accent Darker 50%'!B3", nameof(XLSparklineTheme.Accent3Darker50))]
        [TestCase("'Accent Darker 50%'!B4", nameof(XLSparklineTheme.Accent4Darker50))]
        [TestCase("'Accent Darker 50%'!B5", nameof(XLSparklineTheme.Accent5Darker50))]
        [TestCase("'Accent Darker 50%'!B6", nameof(XLSparklineTheme.Accent6Darker50))]
        [TestCase("'Accent Lighter 40%'!B1", nameof(XLSparklineTheme.Accent1Lighter40))]
        [TestCase("'Accent Lighter 40%'!B2", nameof(XLSparklineTheme.Accent2Lighter40))]
        [TestCase("'Accent Lighter 40%'!B3", nameof(XLSparklineTheme.Accent3Lighter40))]
        [TestCase("'Accent Lighter 40%'!B4", nameof(XLSparklineTheme.Accent4Lighter40))]
        [TestCase("'Accent Lighter 40%'!B5", nameof(XLSparklineTheme.Accent5Lighter40))]
        [TestCase("'Accent Lighter 40%'!B6", nameof(XLSparklineTheme.Accent6Lighter40))]
        [TestCase("Dark!B1", nameof(XLSparklineTheme.Dark1))]
        [TestCase("Dark!B2", nameof(XLSparklineTheme.Dark2))]
        [TestCase("Dark!B3", nameof(XLSparklineTheme.Dark3))]
        [TestCase("Dark!B4", nameof(XLSparklineTheme.Dark4))]
        [TestCase("Dark!B5", nameof(XLSparklineTheme.Dark5))]
        [TestCase("Dark!B6", nameof(XLSparklineTheme.Dark6))]
        [TestCase("Colorful!B1", nameof(XLSparklineTheme.Colorful1))]
        [TestCase("Colorful!B2", nameof(XLSparklineTheme.Colorful2))]
        [TestCase("Colorful!B3", nameof(XLSparklineTheme.Colorful3))]
        [TestCase("Colorful!B4", nameof(XLSparklineTheme.Colorful4))]
        [TestCase("Colorful!B5", nameof(XLSparklineTheme.Colorful5))]
        [TestCase("Colorful!B6", nameof(XLSparklineTheme.Colorful6))]
        public void SparklineThemesAreIdenticalToExcel(string cellAddress, string expectedThemeName)
        {
            using (var ms = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx")))
            using (var wb = new XLWorkbook(ms))
            {
                var expectedStyle = GetThemeByName(expectedThemeName);
                var actualStyle = wb.Cell(cellAddress).Sparkline.SparklineGroup.Style;

                Assert.AreEqual(expectedStyle, actualStyle);
            }

            IXLSparklineStyle GetThemeByName(string themeName)
            {
                var themes = typeof(XLSparklineTheme);
                var prop = themes.GetProperty(themeName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                return prop.GetValue(null, null) as IXLSparklineStyle;
            }
        }

        [Test]
        public void DeletedSparklinesRemovedFromFile()
        {
            using (var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx")))
            using (var output = new MemoryStream())
            {
                using (var wb = new XLWorkbook(input))
                {
                    wb.Worksheet(1).SparklineGroups.RemoveAll();
                    wb.Worksheet(2).SparklineGroups.Remove(wb.Worksheet(2).Cell("B1"));
                    wb.Worksheet(3).SparklineGroups.Remove(wb.Worksheet(3).Range("B2:B6"));
                    wb.Worksheet(4).SparklineGroups.Remove(wb.Worksheet(4).SparklineGroups.First());

                    wb.SaveAs(output);
                }

                using (var wb = new XLWorkbook(output))
                {
                    Assert.AreEqual(0, wb.Worksheet(1).SparklineGroups.Count());
                    Assert.AreEqual(5, wb.Worksheet(2).SparklineGroups.Count());
                    Assert.AreEqual(1, wb.Worksheet(3).SparklineGroups.Count());
                    Assert.AreEqual(5, wb.Worksheet(4).SparklineGroups.Count());
                    Assert.AreEqual(6, wb.Worksheet(5).SparklineGroups.Count());
                    Assert.AreEqual(6, wb.Worksheet(6).SparklineGroups.Count());
                }
            }
        }

        [Test]
        public void EmptySparklineGroupsSkippedOnSaving()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet 1");
                    var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

                    group.RemoveAll();

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual(0, wb.Worksheets.First().SparklineGroups.Count());
                }
            }
        }

        [Test]
        public void CanSaveAndLoadSparklineWithInvalidRange()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws1 = wb.AddWorksheet("Sheet 1");
                    var ws2 = wb.AddWorksheet("Sheet 2");

                    ws1.SparklineGroups.Add("A1:A3", "'Sheet 2'!B1:F3");
                    ws1.SparklineGroups.Add("A4:A6", "B4:F6")
                        .SetDateRange(ws2.Range("A1:E1"));

                    ws2.Delete();
                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.Single();

                    Assert.AreEqual(2, ws.SparklineGroups.Count());
                    Assert.IsFalse(ws.Cell("A2").Sparkline.IsValid);
                    Assert.AreEqual("B5:F5", ws.Cell("A5").Sparkline.SourceData.RangeAddress.ToString());
                    Assert.IsNull(ws.Cell("A5").Sparkline.SparklineGroup.DateRange);
                }
            }
        }

        #endregion Load and save sparkline groups

        #region Change sparkline groups

        [Test]
        public void SetManualMinChangesAxisTypeToCustom()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var axis = ws.SparklineGroups.Add("A1:A2", "B1:Z2")
                .VerticalAxis
                .SetMinAxisType(XLSparklineAxisMinMax.SameForAll);

            axis.ManualMin = 100;

            Assert.AreEqual(100, axis.ManualMin, XLHelper.Epsilon);
            Assert.AreEqual(XLSparklineAxisMinMax.Custom, axis.MinAxisType);
        }

        [Test]
        public void SetManualMaxChangesAxisTypeToCustom()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var axis = ws.SparklineGroups.Add("A1:A2", "B1:Z2")
                .VerticalAxis
                .SetMaxAxisType(XLSparklineAxisMinMax.SameForAll);

            axis.ManualMax = 100;

            Assert.AreEqual(100, axis.ManualMax, XLHelper.Epsilon);
            Assert.AreEqual(XLSparklineAxisMinMax.Custom, axis.MaxAxisType);
        }

        [TestCase(XLSparklineAxisMinMax.Custom, 100)]
        [TestCase(XLSparklineAxisMinMax.SameForAll, null)]
        [TestCase(XLSparklineAxisMinMax.Automatic, null)]
        public void SetAxisTypeToNonCustomSetsManualMinToNull(XLSparklineAxisMinMax axisType, double? expectedManualMin)
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var axis = ws.SparklineGroups.Add("A1", "B1:Z1")
                .VerticalAxis
                .SetManualMin(100);

            axis.MinAxisType = axisType;

            if (expectedManualMin.HasValue)
                Assert.AreEqual(expectedManualMin.Value, axis.ManualMin.Value, XLHelper.Epsilon);
            else
                Assert.IsNull(axis.ManualMin);
        }

        [TestCase(XLSparklineAxisMinMax.Custom, 100)]
        [TestCase(XLSparklineAxisMinMax.SameForAll, null)]
        [TestCase(XLSparklineAxisMinMax.Automatic, null)]
        public void SetAxisTypeToNonCustomSetsManualMaxToNull(XLSparklineAxisMinMax axisType, double? expectedManualMax)
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var axis = ws.SparklineGroups.Add("A1", "B1:Z1")
                .VerticalAxis
                .SetManualMax(100);

            axis.MaxAxisType = axisType;

            if (expectedManualMax.HasValue)
                Assert.AreEqual(expectedManualMax.Value, axis.ManualMax.Value, XLHelper.Epsilon);
            else
                Assert.IsNull(axis.ManualMax);
        }

        [Test]
        public void SetDateRangeChangesAxisType()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

            group.DateRange = ws.Range("B3:Z3");

            Assert.IsTrue(group.HorizontalAxis.DateAxis);
        }

        [Test]
        public void SetDateRangeToNullChangesAxisType()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");
            group.DateRange = ws.Range("B3:Z3");

            group.DateRange = null;

            Assert.IsFalse(group.HorizontalAxis.DateAxis);
        }

        [Test]
        public void CannotSetNonLinearDateRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

            TestDelegate action = () => group.DateRange = ws.Range("B3:Z4");

            Assert.Throws<ArgumentException>(action);
        }

        #endregion Change sparkline groups

        #region Copy sparkline groups

        [Test]
        public void CopyCellToSameWorksheetCopiesSparkline()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");
            ws.SparklineGroups.Add("A1:A3", "B1:F3");
            var target = ws.Cell("D4");

            ws.Cell("A2").CopyTo(target);

            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.IsTrue(target.HasSparkline);
            Assert.AreSame(ws.Cell("A2").Sparkline.SparklineGroup, target.Sparkline.SparklineGroup);
            Assert.AreEqual("E4:I4", target.Sparkline.SourceData.RangeAddress.ToString());
        }

        [Test]
        public void CopyCellToDifferentWorksheetCopiesSparklineGroup()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");
            var ws3 = wb.AddWorksheet("Sheet 3");
            ws1.SparklineGroups.Add("A1:A3", "B1:F3");
            ws1.SparklineGroups.Add("A4:A6", "'Sheet 3'!B4:F6");
            var target1 = ws2.Cell("D4");
            var target2 = ws2.Cell("D5");

            ws1.Cell("A2").CopyTo(target1);
            ws1.Cell("A5").CopyTo(target2);

            Assert.AreEqual(2, ws1.SparklineGroups.Count());
            Assert.AreEqual(2, ws2.SparklineGroups.Count());
            Assert.IsTrue(target1.HasSparkline);
            Assert.IsTrue(target2.HasSparkline);
            Assert.AreEqual("'Sheet 2'!E4:I4", target1.Sparkline.SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 3'!E5:I5", target2.Sparkline.SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true));
        }

        [Test]
        public void CopySparklineIfDateRangeOnSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");
            var group = ws1.SparklineGroups.Add("A1:A3", "B1:F3");
            group.SetDateRange(ws1.Range("A4:E4"));
            var target = ws2.Cell("D4");

            ws1.Cell("A2").CopyTo(target);

            Assert.AreEqual(1, ws1.SparklineGroups.Count());
            Assert.AreEqual(1, ws2.SparklineGroups.Count());
            Assert.IsTrue(target.HasSparkline);
            Assert.AreEqual("'Sheet 2'!D6:H6", target.Sparkline.SparklineGroup.DateRange.RangeAddress.ToString(XLReferenceStyle.A1, true));
        }

        [Test]
        public void CopySparklineIfDateRangeSourceOnDifferentWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");
            var ws3 = wb.AddWorksheet("Sheet 3");
            var group = ws1.SparklineGroups.Add("A1:A3", "B1:F3");
            group.SetDateRange(ws3.Range("A4:E4"));
            var target = ws2.Cell("D4");

            ws1.Cell("A2").CopyTo(target);

            Assert.AreEqual(1, ws1.SparklineGroups.Count());
            Assert.AreEqual(1, ws2.SparklineGroups.Count());
            Assert.IsTrue(target.HasSparkline);
            Assert.AreEqual("'Sheet 3'!D6:H6", target.Sparkline.SparklineGroup.DateRange.RangeAddress.ToString(XLReferenceStyle.A1, true));
        }

        #endregion Copy sparkline groups

        #region Test Examples

        [Test]
        public void CreateSampleSparklines()
        {
            TestHelper.RunTestExample<SampleSparklines>(@"Sparklines\SampleSparklines.xlsx");
        }

        #endregion Test Examples
    }
}
