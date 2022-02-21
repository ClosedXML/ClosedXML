using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Excel.DataValidations
{
    [TestFixture]
    public class DataValidationTests
    {
        [Test]
        public void Validation_1()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Data Validation Issue");
            IXLCell cell = ws.Cell("E1");
            cell.SetValue("Value 1");
            cell = cell.CellBelow();
            cell.SetValue("Value 2");
            cell = cell.CellBelow();
            cell.SetValue("Value 3");
            cell = cell.CellBelow();
            cell.SetValue("Value 4");

            ws.Cell("A1").SetValue("Cell below has Validation Only.");
            cell = ws.Cell("A2");
            cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));

            ws.Cell("B1").SetValue("Cell below has Validation with a title.");
            cell = ws.Cell("B2");
            cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
            cell.GetDataValidation().InputTitle = "Title for B2";

            Assert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
            Assert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
            Assert.AreEqual("Title for B2", cell.GetDataValidation().InputTitle);

            ws.Cell("C1").SetValue("Cell below has Validation with a message.");
            cell = ws.Cell("C2");
            cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
            cell.GetDataValidation().InputMessage = "Message for C2";

            Assert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
            Assert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
            Assert.AreEqual("Message for C2", cell.GetDataValidation().InputMessage);

            ws.Cell("D1").SetValue("Cell below has Validation with title and message.");
            cell = ws.Cell("D2");
            cell.GetDataValidation().List(ws.Range("$E$1:$E$4"));
            cell.GetDataValidation().InputTitle = "Title for D2";
            cell.GetDataValidation().InputMessage = "Message for D2";

            Assert.AreEqual(XLAllowedValues.List, cell.GetDataValidation().AllowedValues);
            Assert.AreEqual("'Data Validation Issue'!$E$1:$E$4", cell.GetDataValidation().Value);
            Assert.AreEqual("Title for D2", cell.GetDataValidation().InputTitle);
            Assert.AreEqual("Message for D2", cell.GetDataValidation().InputMessage);
        }

        [Test]
        public void Validation_2()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").SetValue("A");
            ws.Cell("B1").CreateDataValidation().Custom("Sheet1!A1");

            IXLWorksheet ws2 = wb.AddWorksheet("Sheet2");
            ws2.Cell("A1").SetValue("B");
            ws.Cell("B1").CopyTo(ws2.Cell("B1"));

            Assert.AreEqual("Sheet1!A1", ws2.Cell("B1").GetDataValidation().Value);
        }

        [Test, Ignore("Wait for proper formula shifting (#686)")]
        public void Validation_3()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").SetValue("A");
            ws.Cell("B1").CreateDataValidation().Custom("A1");
            ws.FirstRow().InsertRowsAbove(1);

            Assert.AreEqual("A2", ws.Cell("B2").GetDataValidation().Value);
        }

        [Test]
        public void Validation_4()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").SetValue("A");
            ws.Cell("B1").CreateDataValidation().Custom("A1");
            ws.Cell("B1").CopyTo(ws.Cell("B2"));
            Assert.AreEqual("A2", ws.Cell("B2").GetDataValidation().Value);
        }

        [Test, Ignore("Wait for proper formula shifting (#686)")]
        public void Validation_5()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").SetValue("A");
            ws.Cell("B1").CreateDataValidation().Custom("A1");
            ws.FirstColumn().InsertColumnsBefore(1);

            Assert.AreEqual("B1", ws.Cell("C1").GetDataValidation().Value);
        }

        [Test]
        public void Validation_6()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").SetValue("A");
            ws.Cell("B1").CreateDataValidation().Custom("A1");
            ws.Cell("B1").CopyTo(ws.Cell("C1"));
            Assert.AreEqual("B1", ws.Cell("C1").GetDataValidation().Value);
        }

        [Test]
        public void Validation_persists_on_Cell_DataValidation()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("People");

            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A")
                .CellBelow().SetValue("B");

            IXLTable table = ws.RangeUsed().CreateTable();

            IXLDataValidation dv = table.DataRange.CreateDataValidation();
            dv.ErrorTitle = "Error";

            Assert.AreEqual("Error", table.DataRange.FirstCell().GetDataValidation().ErrorTitle);
        }

        [Test]
        public void Validation_persists_on_Worksheet_DataValidations()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("People");

            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("A");

            IXLTable table = ws.RangeUsed().CreateTable();

            IXLDataValidation dv = table.DataRange.CreateDataValidation();
            dv.ErrorTitle = "Error";

            Assert.AreEqual("Error", ws.DataValidations.Single().ErrorTitle);
        }

        [Test]
        [TestCase("A1:C3", 5, false, "A1:C3")]
        [TestCase("A1:C3", 2, false, "A1:C4")]
        [TestCase("A1:C3", 1, false, "A2:C4")]
        [TestCase("A1:C3", 5, true, "A1:C3")]
        [TestCase("A1:C3", 2, true, "A1:C4")]
        [TestCase("A1:C3", 1, true, "A2:C4")]
        public void DataValidationShiftedOnRowInsert(string initialAddress, int rowNum, bool setValue, string expectedAddress)
        {
            //Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("DataValidation");
            var validation = ws.Range(initialAddress).CreateDataValidation();
            validation.WholeNumber.Between(0, 100);
            if (setValue)
                ws.Range(initialAddress).Value = 50;

            //Act
            ws.Row(rowNum).InsertRowsAbove(1);

            //Assert
            Assert.AreEqual(1, ws.DataValidations.Count());
            Assert.AreEqual(1, ws.DataValidations.First().Ranges.Count());
            Assert.AreEqual(expectedAddress, ws.DataValidations.First().Ranges.First().RangeAddress.ToString());
        }

        [Test]
        [TestCase("A1:C3", 5, false, "A1:C3")]
        [TestCase("A1:C3", 2, false, "A1:D3")]
        [TestCase("A1:C3", 1, false, "B1:D3")]
        [TestCase("A1:C3", 5, true, "A1:C3")]
        [TestCase("A1:C3", 2, true, "A1:D3")]
        [TestCase("A1:C3", 1, true, "B1:D3")]
        public void DataValidationShiftedOnColumnInsert(string initialAddress, int columnNum, bool setValue, string expectedAddress)
        {
            //Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("DataValidation");
            var validation = ws.Range(initialAddress).CreateDataValidation();
            validation.WholeNumber.Between(0, 100);
            if (setValue)
                ws.Range(initialAddress).Value = 50;

            //Act
            ws.Column(columnNum).InsertColumnsBefore(1);

            //Assert
            Assert.AreEqual(1, ws.DataValidations.Count());
            Assert.AreEqual(1, ws.DataValidations.First().Ranges.Count());
            Assert.AreEqual(expectedAddress, ws.DataValidations.First().Ranges.First().RangeAddress.ToString());
        }

        [Test]
        public void DataValidationClearSplitsRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("DataValidation");
                var validation = ws.Range("A1:C3").CreateDataValidation();
                validation.WholeNumber.Between(0, 100);

                //Act
                ws.Cell("B2").Clear(XLClearOptions.DataValidation);

                //Assert
                Assert.IsFalse(ws.Cell("B2").HasDataValidation);
                Assert.IsTrue(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2").All(c => c.HasDataValidation));
            }
        }

        [Test]
        public void NewDataValidationSplitsRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("DataValidation");
                var validation = ws.Range("A1:C3").CreateDataValidation();
                validation.WholeNumber.Between(10, 100);

                //Act
                ws.Cell("B2").CreateDataValidation().WholeNumber.Between(-100, -0);

                //Assert
                Assert.AreEqual("-100", ws.Cell("B2").GetDataValidation().MinValue);
                Assert.IsTrue(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2").All(c => c.HasDataValidation));
                Assert.IsTrue(ws.Range("A1:C3").Cells().Where(c => c.Address.ToString() != "B2")
                                .All(c => c.GetDataValidation().MinValue == "10"));
            }
        }

        [Test]
        public void ListLengthOverflow()
        {
            var values = string.Join(",", Enumerable.Range(1, 20)
                .Select(i => Guid.NewGuid().ToString("N")));

            Assert.True(values.Length > 255);

            using (var wb = new XLWorkbook())
            {
                var dv = wb.AddWorksheet("Sheet 1").Cell(1, 1).GetDataValidation();

                Assert.Throws<ArgumentOutOfRangeException>(() => dv.List(values));
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    dv.TextLength.Between(0, 5);
                    dv.MinValue = values;
                });

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    dv.TextLength.Between(0, 5);
                    dv.MaxValue = values;
                });
            }
        }

        [Test]
        public void CannotCreateDataValidationWithoutRange()
        {
            Assert.Throws<ArgumentNullException>(() => new XLDataValidation(null));
        }

        [Test]
        public void DataValidationHasWorksheetAndRangesWhenCreated()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range = ws.Range("A1:A3");

                var dv = new XLDataValidation(range);

                Assert.AreSame(ws, dv.Worksheet);
                Assert.AreSame(range, dv.Ranges.Single());
            }
        }

        [Test]
        public void CanAddRangeFromSameWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var ranges3 = ws.Ranges("D1:D3,F1:F3");
                var dv = new XLDataValidation(range1);

                dv.AddRange(range2);
                dv.AddRanges(ranges3);

                Assert.IsTrue(dv.Ranges.Any(r => r == range1));
                Assert.IsTrue(dv.Ranges.Any(r => r == range2));
                Assert.IsTrue(dv.Ranges.Any(r => r == ranges3.First()));
                Assert.IsTrue(dv.Ranges.Any(r => r == ranges3.Last()));
            }
        }

        [Test]
        public void CanAddRangeFromAnotherWorksheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet();
                var ws2 = wb.AddWorksheet();
                var range1 = ws1.Range("A1:A3");
                var range2 = ws2.Range("C1:C3");
                var dv = new XLDataValidation(range1);

                dv.AddRange(range2);

                Assert.IsTrue(dv.Ranges.Any(r => r != range2 && r.RangeAddress.ToString() == range2.RangeAddress.ToString()));
            }
        }

        [Test]
        public void CanClearRanges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var ranges3 = ws.Ranges("D1:D3,F1:F3");
                var dv = new XLDataValidation(range1);
                dv.AddRange(range2);
                dv.AddRanges(ranges3);

                dv.ClearRanges();

                Assert.IsEmpty(dv.Ranges);
            }
        }

        [Test]
        public void CanRemoveExistingRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");

                var dv = new XLDataValidation(range1);
                dv.AddRange(range2);

                dv.RemoveRange(range1);

                Assert.AreSame(range2, dv.Ranges.Single());
            }
        }

        [Test]
        public void RemovingExistingRangeDoesNoFail()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");

                var dv = new XLDataValidation(range1);

                dv.RemoveRange(range2);
                dv.RemoveRange(null);

                Assert.AreSame(range1, dv.Ranges.Single());
            }
        }

        [Test]
        public void AddRangeFiresEvent()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var dv = new XLDataValidation(range1);

                IXLRange addedRange = null;

                dv.RangeAdded += (s, e) => addedRange = e.Range;

                dv.AddRange(range2);

                Assert.AreSame(range2, addedRange);
            }
        }

        [Test]
        public void AddRangesFiresMultipleEvents()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var ranges = ws.Ranges("D1:D3,F1:F3");
                var dv = new XLDataValidation(range1);

                var addedRanges = new List<IXLRange>();

                dv.RangeAdded += (s, e) => addedRanges.Add(e.Range);

                dv.AddRanges(ranges);

                Assert.AreEqual(2, addedRanges.Count);
            }
        }

        [Test]
        public void RemoveRangeFiresEvent()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var dv = new XLDataValidation(range1);
                dv.AddRange(range2);
                IXLRange removedRange = null;
                dv.RangeRemoved += (s, e) => removedRange = e.Range;

                dv.RemoveRange(range2);

                Assert.AreSame(range2, removedRange);
            }
        }

        [Test]
        public void RemoveNonExistingRangeDoesNotFireEvent()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var dv = new XLDataValidation(range1);

                dv.RangeRemoved += (s, e) => Assert.Fail("Expected not to fire event");

                dv.RemoveRange(range2);
            }
        }

        [Test]
        public void ClearRangesFiresMultipleEvents()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var range1 = ws.Range("A1:A3");
                var range2 = ws.Range("C1:C3");
                var dv = new XLDataValidation(range1);
                dv.AddRange(range2);

                var removedRanges = new List<IXLRange>();

                dv.RangeRemoved += (s, e) => removedRanges.Add(e.Range);

                dv.ClearRanges();

                Assert.AreEqual(2, removedRanges.Count);
            }
        }
    }
}
