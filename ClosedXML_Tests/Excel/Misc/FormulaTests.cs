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
    }
}
