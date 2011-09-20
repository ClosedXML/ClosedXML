using ClosedXML.Excel;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class NamedRangesTests
    {

        [TestMethod]
        public void MovingRanges()
        {
            var wb = new XLWorkbook();

            var sheet1 = wb.Worksheets.Add("Sheet1");
            var sheet2 = wb.Worksheets.Add("Sheet2");

            wb.NamedRanges.Add("wbNamedRange",
                               "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.NamedRanges.Add("sheet1NamedRange",
                               "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            sheet1.Row(1).InsertRowsAbove(2);
            sheet1.Row(1).Delete();
            sheet1.Column(1).InsertColumnsBefore(2);
            sheet1.Column(1).Delete();


            Assert.AreEqual("'Sheet1'!$C$3,'Sheet1'!$C$4:$D$4,Sheet2!$D$3:$D$4,'Sheet1'!$7:$8,'Sheet1'!$G:$H", wb.NamedRanges.First().RefersTo);
            Assert.AreEqual("'Sheet1'!$C$3,'Sheet1'!$C$4:$D$4,Sheet2!$D$3:$D$4,'Sheet1'!$7:$8,'Sheet1'!$G:$H", sheet1.NamedRanges.First().RefersTo);
            Assert.AreEqual("'Sheet1'!B2,Sheet2!A1", sheet2.NamedRanges.First().RefersTo);
        }

    }
}
