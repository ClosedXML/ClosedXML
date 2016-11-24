using ClosedXML.Excel;
using NUnit.Framework;

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
    }
}
