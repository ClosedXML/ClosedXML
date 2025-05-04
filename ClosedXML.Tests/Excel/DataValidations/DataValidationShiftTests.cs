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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("DataValidationShift");
            ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Cells("A1:D7").Value = 1;

            ws.Column(2).InsertColumnsAfter(2);
            var dv = ws.DataValidations.ToArray();

            Assert.That(dv.Length, Is.EqualTo(5));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            Assert.That(dv[1].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A2:D2"));
            Assert.That(dv[2].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A3:E3"));
            Assert.That(dv[3].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("B4:D6"));
            Assert.That(dv[4].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("E7:F7"));
        }

        [Test]
        public void DataValidationShiftedOnRowInsert()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("DataValidationShift");
            ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Cells("A1:G5").Value = 1;

            ws.Row(2).InsertRowsBelow(2);
            var dv = ws.DataValidations.ToArray();

            Assert.That(dv.Length, Is.EqualTo(5));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            Assert.That(dv[1].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("B1:B4"));
            Assert.That(dv[2].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("C1:C5"));
            Assert.That(dv[3].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("D2:F4"));
            Assert.That(dv[4].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("G6:G7"));
        }

        [Test]
        public void DataValidationShiftedOnColumnDelete()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("DataValidationShift");
            ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Cells("A1:D7").Value = 1;

            ws.Column(2).Delete();
            var dv = ws.DataValidations.ToArray();

            Assert.That(dv.Length, Is.EqualTo(4));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            Assert.That(dv[1].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A2:A2"));
            Assert.That(dv[2].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A3:B3"));
            Assert.That(dv[3].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("B7:C7"));
        }

        [Test]
        public void DataValidationShiftedOnRowDelete()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("DataValidationShift");
            ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
            ws.Cells("A1:G5").Value = 1;

            ws.Row(2).Delete();
            var dv = ws.DataValidations.ToArray();

            Assert.That(dv.Length, Is.EqualTo(4));
            Assert.That(dv[0].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            Assert.That(dv[1].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("B1:B1"));
            Assert.That(dv[2].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("C1:C2"));
            Assert.That(dv[3].Ranges.Single().RangeAddress.ToString(), Is.EqualTo("G3:G4"));
        }

        [Test]
        public void DataValidationShiftedTruncateRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("DataValidationShift");
            ws.AsRange().CreateDataValidation().WholeNumber.Between(0, 1);
            var dv = ws.DataValidations.Single();

            ws.Row(2).InsertRowsAbove(1);
            Assert.Multiple(() =>
            {
                Assert.That(dv.Ranges.Single().RangeAddress.IsValid, Is.True);
                Assert.That(dv.Ranges.Single().RangeAddress.ToString(), Is.EqualTo($"1:{XLHelper.MaxRowNumber}"));
            });

            ws.Column(2).InsertColumnsAfter(1);
            Assert.Multiple(() =>
            {
                Assert.That(dv.Ranges.Single().RangeAddress.IsValid, Is.True);
                Assert.That(dv.Ranges.Single().RangeAddress.ToString(), Is.EqualTo($"1:{XLHelper.MaxRowNumber}"));
            });
        }
    }
}
