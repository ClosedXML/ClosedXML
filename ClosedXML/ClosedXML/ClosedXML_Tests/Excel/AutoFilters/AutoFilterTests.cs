using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;

namespace ClosedXML_Tests
{
    [TestClass()]
    public class AutoFilterTests
    {
        [TestMethod()]
        public void AutoFilterSortWhenNotInFirstRow()
        {
            using (var wb = new XLWorkbook())
            {
                using (var ws = wb.Worksheets.Add("Sheet1"))
                {
                    ws.Cell(3, 3).SetValue("Names")
                        .CellBelow().SetValue("Manuel")
                        .CellBelow().SetValue("Carlos")
                        .CellBelow().SetValue("Dominic");
                    ws.RangeUsed().SetAutoFilter().Sort();
                    Assert.AreEqual(ws.Cell(4, 3).GetString(), "Carlos");
                }
                
            }
            
        }

    }
}
