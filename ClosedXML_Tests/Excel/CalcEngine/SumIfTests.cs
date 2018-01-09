using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class SumIfTests
    {
        [TestCase("SUMIF Not Same Columns", "E1", 3)]
        [TestCase("SUMIF Same Columns", "B7", 3)]
        public void CheckCellValues(string sheetName, string cellReference, object expectedValue)
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Functions\TestSumIf.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet(sheetName);
                var value = ws.Cell(cellReference).Value;
                Assert.AreEqual(expectedValue, value);                
            }
        }

        
    }
}
