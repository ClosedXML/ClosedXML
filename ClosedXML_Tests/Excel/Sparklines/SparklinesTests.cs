using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.Sparklines
{
    [TestFixture]
    public class SparklinesTests
    {
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
        public void CannotAddSparklineForWhenRangesHaveDifferentWidths()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:C1"), ws.Range("A3:D4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("locationRange and sourceDataRange must have the same width", message);
        }

        [Test]
        public void CannotAddSparklineForWhenRangesHaveDifferentHeights()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            TestDelegate action = () => ws.SparklineGroups.Add(ws.Range("A1:A3"), ws.Range("B1:B4"));

            var message = Assert.Throws<ArgumentException>(action).Message;
            Assert.AreEqual("locationRange and sourceDataRange must have the same height", message);
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
    }
}
