using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class FormulaTests
    {
        [Test]
        public void CopyFormula()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").FormulaA1 = "B1";
            ws.Cell("A1").CopyTo("A2");
            Assert.AreEqual("B2", ws.Cell("A2").FormulaA1);
        }

        [Test]
        public void CopyFormula2()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").FormulaA1 = "A2-1";
                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("R[1]C-1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("R[1]C-1", ws.Cell("B1").FormulaR1C1);
                Assert.AreEqual("B2-1", ws.Cell("B1").FormulaA1);

                ws.Cell("A1").FormulaA1 = "B1+1";
                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("RC[1]+1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("RC[1]+1", ws.Cell("A2").FormulaR1C1);
                Assert.AreEqual("B2+1", ws.Cell("A2").FormulaA1);
            }
        }

        [Test]
        public void CopyFormulaWithSheetNameThatResemblesFormula()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("S10 Data");
                ws.Cell("A1").Value = "Some value";

                ws = wb.Worksheets.Add("Summary");
                ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
                Assert.AreEqual("Some value", ws.Cell("A1").Value);

                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("'S10 Data'!A2", ws.Cell("A2").FormulaA1);

                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("'S10 Data'!B1", ws.Cell("B1").FormulaA1);
            }
        }

        [Test]
        public void DateAgainstStringComparison()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = new DateTime(2016, 1, 1);
                ws.Cell("A1").DataType = XLCellValues.DateTime;

                ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
                var actual = ws.Cell("A2").Value;
                Assert.AreEqual(actual, "B");

                ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
                actual = ws.Cell("A3").Value;
                Assert.AreEqual(actual, "B");
            }
        }
    }
}
