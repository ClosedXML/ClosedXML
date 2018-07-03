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

            var message = Assert.Throws<ArgumentException>(action = ).Message;
            Assert.AreEqual("locationRange and sourceDataRange must have the same height", message);
        }
    }
}
