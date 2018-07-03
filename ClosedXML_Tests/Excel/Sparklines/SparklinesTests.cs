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
        public void CanAddSparklines()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet 1");

            ws.SparklineGroups.Add("A1", "B1:E1");
            
            Assert.AreEqual(1, ws.SparklineGroups.Count());
            Assert.AreEqual(1, ws.SparklineGroups.Single().Count());
            Assert.AreEqual("A1", ws.SparklineGroups.Single().Single().Location.Address.ToString());
            Assert.AreEqual("B1:E1", ws.SparklineGroups.Single().Single().SourceData.RangeAddress.ToString());
        }
    }
}
