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
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var ws2 = wb.AddWorksheet();

            var dv1 = ws1.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";

            var dv2 = ws2.DataValidations.Add(dv1);

            Assert.Multiple(() =>
            {
                Assert.That(ws1.DataValidations.Count(), Is.EqualTo(1));
                Assert.That(ws2.DataValidations.Count(), Is.EqualTo(1));
            });

            Assert.Multiple(() =>
            {
                Assert.That(dv2, Is.Not.SameAs(dv1));

                Assert.That(dv1.Ranges.Single().Worksheet, Is.SameAs(ws1));
            });
            Assert.That(dv2.Ranges.Single().Worksheet, Is.SameAs(ws2));
        }

        [TestCase("A1:A1", true)]
        [TestCase("A1:A3", true)]
        [TestCase("A1:A4", false)]
        [TestCase("C2:C2", true)]
        [TestCase("C1:C3", true)]
        [TestCase("A1:C3", false)]
        public void CanFindDataValidationForRange(string searchAddress, bool expectedResult)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var dv = ws.Range("A1:A3").CreateDataValidation();
            dv.MinValue = "100";
            dv.AddRange(ws.Range("C1:C3"));

            var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

            var actualResult = ws.DataValidations.TryGet(address, out var foundDv);
            Assert.That(actualResult, Is.EqualTo(expectedResult));
            if (expectedResult)
                Assert.That(foundDv, Is.SameAs(dv));
            else
                Assert.That(foundDv, Is.Null);
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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var dv1 = ws.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";
            dv1.AddRange(ws.Range("C1:C3"));

            var dv2 = ws.Range("E4:G6").CreateDataValidation();
            dv2.MinValue = "200";

            var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

            var actualResult = ws.DataValidations.GetAllInRange(address);

            Assert.That(actualResult.Count(), Is.EqualTo(expectedCount));
        }

        [Test]
        public void AddDataValidationSplitsExistingRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var dv1 = ws.Ranges("B2:G7,C11:C13").CreateDataValidation();
            dv1.MinValue = "100";

            var dv2 = ws.Range("E4:G6").CreateDataValidation();
            dv2.MinValue = "100";

            Assert.Multiple(() =>
            {
                Assert.That(dv1.Ranges.Count(), Is.EqualTo(4));
                Assert.That(string.Join(",", dv1.Ranges.Select(r => r.RangeAddress.ToString())),
                    Is.EqualTo("B2:G3,B4:D6,B7:G7,C11:C13"));
            });
        }

        [Test]
        public void RemovedRangeExcludedFromIndex()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var dv = ws.Range("A1:A3").CreateDataValidation();
            dv.MinValue = "100";
            var range = ws.Range("C1:C3");
            dv.AddRange(range);

            dv.RemoveRange(range);

            var actualResult = ws.DataValidations.TryGet(range.RangeAddress, out var foundDv);
            Assert.That(actualResult, Is.False);
            Assert.That(foundDv, Is.Null);
        }

        [Test]
        public void ConsolidatedDataValidationsAreUnsubscribed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var dv1 = ws.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";
            var dv2 = ws.Range("B1:B3").CreateDataValidation();
            dv2.MinValue = "100";

            (ws.DataValidations as XLDataValidations).Consolidate();
            dv1.AddRange(ws.Range("C1:C3"));
            dv2.AddRange(ws.Range("D1:D3"));

            var consolidatedDv = ws.DataValidations.Single();
            Assert.Multiple(() =>
            {
                Assert.That(consolidatedDv, Is.SameAs(dv1));
                Assert.That(ws.Cell("C1").HasDataValidation, Is.True);
            });
            Assert.False(ws.Cell("D1").HasDataValidation);
        }
    }
}
