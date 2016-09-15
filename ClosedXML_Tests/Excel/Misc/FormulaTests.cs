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
    }
}