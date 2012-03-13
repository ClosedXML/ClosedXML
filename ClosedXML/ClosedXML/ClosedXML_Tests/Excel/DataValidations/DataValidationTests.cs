using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.DataValidations
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class DataValidationTests
    {
        [TestMethod]
        public void Validation_persists_on_Worksheet_DataValidations()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("People");

            ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A");

            var table = ws.RangeUsed().CreateTable();

            var dv = table.DataRange.SetDataValidation();
            dv.ErrorTitle = "Error";

            Assert.AreEqual("Error", ws.DataValidations.Single().ErrorTitle);
        }

        [TestMethod]
        public void Validation_persists_on_Cell_DataValidation()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("People");

            ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B");

            var table = ws.RangeUsed().CreateTable();

            var dv = table.DataRange.SetDataValidation();
            dv.ErrorTitle = "Error";

            Assert.AreEqual("Error", table.DataRange.FirstCell().DataValidation.ErrorTitle);
        }

        [TestMethod]
        public void Validation_1()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data Validation Issue");
            var cell = ws.Cell("E1");
            cell.SetValue("Value 1");
            cell = cell.CellBelow();
            cell.SetValue("Value 2");
            cell = cell.CellBelow();
            cell.SetValue("Value 3");
            cell = cell.CellBelow();
            cell.SetValue("Value 4");
            cell = cell.CellBelow();

            ws.Cell("A1").SetValue("Cell below has Validation Only.");
            cell = ws.Cell("A2");
            cell.DataValidation.List(ws.Range("$E$1:$E$4"));

            ws.Cell("B1").SetValue("Cell below has Validation with a title.");
            cell = ws.Cell("B2");
            cell.DataValidation.List(ws.Range("$E$1:$E$4"));
            cell.DataValidation.InputTitle = "Title for B2";

            Assert.AreEqual(cell.DataValidation.AllowedValues, XLAllowedValues.List);
            Assert.AreEqual(cell.DataValidation.Value, "'Data Validation Issue'!$E$1:$E$4");
            Assert.AreEqual(cell.DataValidation.InputTitle, "Title for B2");


            ws.Cell("C1").SetValue("Cell below has Validation with a message.");
            cell = ws.Cell("C2");
            cell.DataValidation.List(ws.Range("$E$1:$E$4"));
            cell.DataValidation.InputMessage = "Message for C2";

            Assert.AreEqual(cell.DataValidation.AllowedValues, XLAllowedValues.List);
            Assert.AreEqual(cell.DataValidation.Value, "'Data Validation Issue'!$E$1:$E$4");
            Assert.AreEqual(cell.DataValidation.InputMessage, "Message for C2");

            ws.Cell("D1").SetValue("Cell below has Validation with title and message.");
            cell = ws.Cell("D2");
            cell.DataValidation.List(ws.Range("$E$1:$E$4"));
            cell.DataValidation.InputTitle = "Title for D2";
            cell.DataValidation.InputMessage = "Message for D2";

            Assert.AreEqual(cell.DataValidation.AllowedValues, XLAllowedValues.List);
            Assert.AreEqual(cell.DataValidation.Value, "'Data Validation Issue'!$E$1:$E$4");
            Assert.AreEqual(cell.DataValidation.InputTitle, "Title for D2");
            Assert.AreEqual(cell.DataValidation.InputMessage, "Message for D2");
        }
    }
}
