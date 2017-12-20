using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel.Workbook
{
    [TestFixture]
    public class XLWorkbookTests
    {
        [Test]
        //When using default constructor like: new Workbook()
        public void ReturnsDefaultToStringWhenStreamAndFileAreNull()
        {
            using (var wb = new XLWorkbook())
            {
                var result = wb.ToString();
            }
        }
    }
}
