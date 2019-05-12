using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.DataValidations
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

                var dv1 = ws1.Range("A1:A3").SetDataValidation();
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
                var dv = ws.Range("A1:A3").SetDataValidation();
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
                var dv1 = ws.Range("A1:A3").SetDataValidation();
                dv1.MinValue = "100";
                dv1.AddRange(ws.Range("C1:C3"));

                var dv2 = ws.Range("E4:G6").SetDataValidation();
                dv2.MinValue = "100";

                var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

                var actualResult = ws.DataValidations.GetAllInRange(address);

                Assert.AreEqual(expectedCount, actualResult.Count());
            }
        }

        [Test]
        public void AddDataValidationSplitsExistingRanges()
        {

        }
    }
}
