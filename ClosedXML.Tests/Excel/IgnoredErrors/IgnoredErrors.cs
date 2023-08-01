using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.IgnoredErrors
{
    [TestFixture]
    public class IgnoredErrors
    {
        //Check, whether mapping for all implemented ignored errors types works in both directions
        [Test]
        public void TestMappingToOpenXml()
        {
            var wb = new XLWorkbook();

            var ws1 = wb.AddWorksheet("T1");

            var types = Enum.GetValues(typeof(XLIgnoredErrorType)).Cast<XLIgnoredErrorType>();
            foreach (var type in types)
            {
                ws1.IgnoredErrors.Add(type, ws1.Range(1, 1, 1, 1));
            }

            var ignoredErrors = XLIgnoredErrorOpenXmlMapper.GetOpenXmlIgnoredErrors(ws1);
            Assert.AreEqual(types.Count(), ignoredErrors.Count());
        }

        //Check save, load and copy worksheet of ignored errors
        [Test]
        public void TestSaveLoadCopy()
        {
            TestHelper.RunTestExample<ClosedXML.Examples.IgnoredErrors.IgnoredErrors>(@"IgnoredErrors\IgnoredErrors.xlsx");
        }
    }
}
