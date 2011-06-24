using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.Misc
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class CopyContentsTest
    {
        

        [TestMethod]
        public void TestRowCopyContents()
        {
            var workbook = new XLWorkbook();
            var originalSheet = workbook.Worksheets.Add("original");
            var copyRowSheet = workbook.Worksheets.Add("copy row");
            var copyRowAsRangeSheet = workbook.Worksheets.Add("copy row as range");
            var copyRangeSheet = workbook.Worksheets.Add("copy range");

            originalSheet.Cell("A2").SetValue("test value");
            originalSheet.Range("A2:E2").Merge();

            
            {
                var originalRange = originalSheet.Range("A2:E2");
                var destinationRange = copyRangeSheet.Range("A2:E2");
                originalRange.CopyTo(destinationRange);
            }

            CopyRowAsRange(originalSheet, 2, copyRowAsRangeSheet, 3);

            {
                var originalRow = originalSheet.Row(2);
                var destinationRow = copyRowSheet.Row(2);
                originalRow.CopyTo(destinationRow);
            }
            TestHelper.SaveWorkbook(workbook, "CopyRowContents.xlsx");
        }

        private static void CopyRowAsRange(IXLWorksheet originalSheet,  int originalRowNumber, IXLWorksheet destSheet, int destRowNumber)
        {
            {
                var destinationRow = destSheet.Row(destRowNumber);
                destinationRow.Clear();

                var originalRow = originalSheet.Row(originalRowNumber);
                int columnNumber= originalRow.LastCellUsed(true).Address.ColumnNumber;

                var originalRange= originalSheet.Range(originalRowNumber, 1, originalRowNumber, columnNumber);
                var destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
                originalRange.CopyTo(destRange);
            }
        }
    }
}
