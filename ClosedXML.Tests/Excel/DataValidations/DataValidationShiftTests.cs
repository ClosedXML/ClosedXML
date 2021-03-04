using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.DataValidations
{
    [TestFixture]
    public class DataValidationShiftTests
    {
        [Test]
        public void DataValidationShiftedOnColumnInsert()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("DataValidationShift");
                ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Cells("A1:D7").Value = 1;

                ws.Column(2).InsertColumnsAfter(2);
                var dv = ws.DataValidations.ToArray();

                Assert.AreEqual(5, dv.Length);
                Assert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("A2:D2", dv[1].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("A3:E3", dv[2].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("B4:D6", dv[3].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("E7:F7", dv[4].Ranges.Single().RangeAddress.ToString());
            }
        }

        [Test]
        public void DataValidationShiftedOnRowInsert()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("DataValidationShift");
                ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Cells("A1:G5").Value = 1;

                ws.Row(2).InsertRowsBelow(2);
                var dv = ws.DataValidations.ToArray();

                Assert.AreEqual(5, dv.Length);
                Assert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("B1:B4", dv[1].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("C1:C5", dv[2].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("D2:F4", dv[3].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("G6:G7", dv[4].Ranges.Single().RangeAddress.ToString());
            }
        }

        [Test]
        public void DataValidationShiftedOnColumnDelete()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("DataValidationShift");
                ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Cells("A1:D7").Value = 1;

                ws.Column(2).Delete();
                var dv = ws.DataValidations.ToArray();

                Assert.AreEqual(4, dv.Length);
                Assert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("A2:A2", dv[1].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("A3:B3", dv[2].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("B7:C7", dv[3].Ranges.Single().RangeAddress.ToString());
            }
        }

        [Test]
        public void DataValidationShiftedOnRowDelete()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("DataValidationShift");
                ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
                ws.Cells("A1:G5").Value = 1;

                ws.Row(2).Delete();
                var dv = ws.DataValidations.ToArray();

                Assert.AreEqual(4, dv.Length);
                Assert.AreEqual("A1:A1", dv[0].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("B1:B1", dv[1].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("C1:C2", dv[2].Ranges.Single().RangeAddress.ToString());
                Assert.AreEqual("G3:G4", dv[3].Ranges.Single().RangeAddress.ToString());
            }
        }

        [Test]
        public void DataValidationShiftedTruncateRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("DataValidationShift");
                ws.AsRange().CreateDataValidation().WholeNumber.Between(0, 1);
                var dv = ws.DataValidations.Single();

                ws.Row(2).InsertRowsAbove(1);
                Assert.IsTrue(dv.Ranges.Single().RangeAddress.IsValid);
                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", dv.Ranges.Single().RangeAddress.ToString());

                ws.Column(2).InsertColumnsAfter(1);
                Assert.IsTrue(dv.Ranges.Single().RangeAddress.IsValid);
                Assert.AreEqual($"1:{XLHelper.MaxRowNumber}", dv.Ranges.Single().RangeAddress.ToString());
            }
        }
    }
}
