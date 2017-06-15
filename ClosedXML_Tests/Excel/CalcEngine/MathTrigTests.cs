using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ClosedXML.Excel;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestClass]
    public class MathTrigTests
    {
        [TestMethod]
        public void MathTrig_RoundShouldRoundAwayFromZero()
        {
            double cellAmount = 31.565d;
            double roundedAmount = Math.Round(cellAmount, 2, MidpointRounding.AwayFromZero);

            using (XLWorkbook testBook = new XLWorkbook(@"D:\Test.xlsx"))
            {
                IXLWorksheet testSheet;
                testBook.TryGetWorksheet("Sheet1", out testSheet);

                double actualCellAmount = testSheet.Cell(2, 2).GetDouble();
                Assert.AreEqual(roundedAmount, actualCellAmount);
            }
        }
    }
}
