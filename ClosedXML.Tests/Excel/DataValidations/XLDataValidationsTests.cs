using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.DataValidations
{
    public class XLDataValidationsTests
    {
        [Test]
        public void CannotCreateWithoutWorksheet()
        {
            Assert.Throws<ArgumentNullException>(() => new XLDataValidations(null));
        }

        [Test]
        public void AddedRangesAreTransferredToTargetSheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var ws2 = wb.AddWorksheet();

                var dv1 = ws1.Range("A1:A3").CreateDataValidation();
                dv1.MinValue = "100";

                var dv2 = ws2.DataValidations.Add(dv1);

                Assert.AreEqual(1, ws1.DataValidations.Count());
                Assert.AreEqual(1, ws2.DataValidations.Count());

                Assert.AreNotSame(dv1, dv2);

                Assert.AreSame(ws1, dv1.Ranges.Single().Worksheet);
                Assert.AreSame(ws2, dv2.Ranges.Single().Worksheet);
            }
        }

        [TestCase("A1:A1", true)]
        [TestCase("A1:A3", true)]
        [TestCase("A1:A4", false)]
        [TestCase("C2:C2", true)]
        [TestCase("C1:C3", true)]
        [TestCase("A1:C3", false)]
        public void CanFindDataValidationForRange(string searchAddress, bool expectedResult)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var dv = ws.Range("A1:A3").CreateDataValidation();
                dv.MinValue = "100";
                dv.AddRange(ws.Range("C1:C3"));

                var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

                var actualResult = ws.DataValidations.TryGet(address, out var foundDv);
                Assert.AreEqual(expectedResult, actualResult);
                if (expectedResult)
                    Assert.AreSame(dv, foundDv);
                else
                    Assert.IsNull(foundDv);
            }
        }

        [TestCase("A1:A1", 1)]
        [TestCase("A1:A3", 1)]
        [TestCase("B1:B4", 0)]
        [TestCase("A1:C3", 1)]
        [TestCase("C2:C3", 1)]
        [TestCase("C2:G6", 2)]
        [TestCase("E2:E3", 0)]
        public void CanGetAllDataValidationsForRange(string searchAddress, int expectedCount)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var dv1 = ws.Range("A1:A3").CreateDataValidation();
                dv1.MinValue = "100";
                dv1.AddRange(ws.Range("C1:C3"));

                var dv2 = ws.Range("E4:G6").CreateDataValidation();
                dv2.MinValue = "200";

                var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

                var actualResult = ws.DataValidations.GetAllInRange(address);

                Assert.AreEqual(expectedCount, actualResult.Count());
            }
        }

        [Test]
        public void AddDataValidationSplitsExistingRanges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var dv1 = ws.Ranges("B2:G7,C11:C13").CreateDataValidation();
                dv1.MinValue = "100";

                var dv2 = ws.Range("E4:G6").CreateDataValidation();
                dv2.MinValue = "100";

                Assert.AreEqual(4, dv1.Ranges.Count());
                Assert.AreEqual("B2:G3,B4:D6,B7:G7,C11:C13",
                    string.Join(",", dv1.Ranges.Select(r => r.RangeAddress.ToString())));
            }
        }

        [Test]
        public void RemovedRangeExcludedFromIndex()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var dv = ws.Range("A1:A3").CreateDataValidation();
                dv.MinValue = "100";
                var range = ws.Range("C1:C3");
                dv.AddRange(range);

                dv.RemoveRange(range);

                var actualResult = ws.DataValidations.TryGet(range.RangeAddress, out var foundDv);
                Assert.IsFalse(actualResult);
                Assert.IsNull(foundDv);
            }
        }

        [Test]
        public void ConsolidatedDataValidationsAreUnsubscribed()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var dv1 = ws.Range("A1:A3").CreateDataValidation();
                dv1.MinValue = "100";
                var dv2 = ws.Range("B1:B3").CreateDataValidation();
                dv2.MinValue = "100";

                (ws.DataValidations as XLDataValidations).Consolidate();
                dv1.AddRange(ws.Range("C1:C3"));
                dv2.AddRange(ws.Range("D1:D3"));

                var consolidatedDv = ws.DataValidations.Single();
                Assert.AreSame(dv1, consolidatedDv);
                Assert.True(ws.Cell("C1").HasDataValidation);
                Assert.False(ws.Cell("D1").HasDataValidation);
            }
        }
    }
}
