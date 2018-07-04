using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Sparklines
{
    [TestFixture]
    public class SparklinesTests
    {
        #region Add sparklines

        [Test]
        public void CanotCreateSparklineGroupsWithoutWorksheet()
        {
            TestDelegate action = () => new XLSparklineGroups(null);
            Assert.Throws<ArgumentNullException>(action);
        }

        [Test]
        public void CanotCreateSparklineGroupWithoutWorksheet()
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
        public void CannotCreateSparklineWithoutSourceData()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var group = new XLSparklineGroup(ws);
            TestDelegate action = () => new XLSparkline(group, ws.FirstCell(), null);
            Assert.Throws<ArgumentNullException>(action);
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
            Assert.AreEqual(99+25, sparklines3.Count());
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
        public void CanChangeSparklineSourceDataDifferentWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");

            ws1.SparklineGroups.Add("A1:A2", "B1:Z2");
            ws1.SparklineGroups.Single().Last().SetSourceData(ws2.Range("D4:D50"));

            Assert.AreEqual(1, ws1.SparklineGroups.Count());
            Assert.AreEqual(2, ws1.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws1.SparklineGroups.Single().First().Location.Address.ToString());
            Assert.AreEqual("A2", ws1.SparklineGroups.Single().Last().Location.Address.ToString());
            Assert.AreEqual("'Sheet 1'!B1:Z1", ws1.SparklineGroups.Single().First().SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 2'!D4:D50", ws1.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true));
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
                        .SetAxisColor(XLColor.AirForceBlue)
                        .SetFirstMarkerColor(XLColor.AliceBlue)
                        .SetHighMarkerColor(XLColor.Alizarin)
                        .SetLastMarkerColor(XLColor.Almond)
                        .SetLowMarkerColor(XLColor.Amaranth)
                        .SetMarkersColor(XLColor.Amber)
                        .SetNegativeColor(XLColor.AmberSaeEce)
                        .SetSeriesColor(XLColor.AmericanRose)
                        .SetLineWeight(5.5)
                        .SetManualMax(6.6)
                        .SetManualMin(1.2)
                        .SetDateAxis(true)
                        .SetDisplayHidden(true)
                        .SetDisplayXAxis(true)
                        .SetFirst(true)
                        .SetHigh(true)
                        .SetLast(true)
                        .SetLow(true)
                        .SetMarkers(true)
                        .SetNegative(true)
                        .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
                        .SetMaxAxisType(XLSparklineAxisMinMax.Custom)
                        .SetMinAxisType(XLSparklineAxisMinMax.Custom)
                        .SetType(XLSparklineType.Stacked);

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

                Assert.AreEqual(XLColor.AirForceBlue, group.AxisColor);
                Assert.AreEqual(XLColor.AliceBlue, group.FirstMarkerColor);
                Assert.AreEqual(XLColor.Alizarin, group.HighMarkerColor);
                Assert.AreEqual(XLColor.Almond, group.LastMarkerColor);
                Assert.AreEqual(XLColor.Amaranth, group.LowMarkerColor);
                Assert.AreEqual(XLColor.Amber, group.MarkersColor);
                Assert.AreEqual(XLColor.AmberSaeEce, group.NegativeColor);
                Assert.AreEqual(XLColor.AmericanRose, group.SeriesColor);

                Assert.AreEqual(5.5, group.LineWeight, XLHelper.Epsilon);
                Assert.AreEqual(6.6, group.ManualMax, XLHelper.Epsilon);
                Assert.AreEqual(1.2, group.ManualMin, XLHelper.Epsilon);

                Assert.IsTrue(group.DisplayHidden);
                Assert.IsTrue(group.DisplayXAxis);
                Assert.IsTrue(group.First);
                Assert.IsTrue(group.High);
                Assert.IsTrue(group.Last);
                Assert.IsTrue(group.Low);
                Assert.IsTrue(group.Markers);
                Assert.IsTrue(group.Negative);

                Assert.AreEqual(XLDisplayBlanksAsValues.Zero, group.DisplayEmptyCellsAs);
                Assert.AreEqual(XLSparklineAxisMinMax.Custom, group.MaxAxisType);
                Assert.AreEqual(XLSparklineAxisMinMax.Custom, group.MinAxisType);
                Assert.AreEqual(XLSparklineType.Stacked, group.Type);
            }
        }

        #endregion Load and save sparkline groups

    }
}
