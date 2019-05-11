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

        [Test]
        public void CanFindDataValidationForRange()
        {

        }

        [Test]
        public void AddDataValidationSplitsExistingRanges()
        {

        }
    }
}
