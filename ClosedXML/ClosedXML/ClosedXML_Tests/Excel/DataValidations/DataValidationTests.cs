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
            .CellBelow().SetValue("A")
            .CellBelow().SetValue("B")
            .CellBelow().SetValue("")
            .CellBelow().SetValue("D");

            var table = ws.RangeUsed().CreateTable();

            var dv = table.DataRange.SetDataValidation();
            dv.ErrorTitle = "Error";

            Assert.AreEqual("Error", ws.DataValidations.Single().ErrorTitle);
        }

    }
}
