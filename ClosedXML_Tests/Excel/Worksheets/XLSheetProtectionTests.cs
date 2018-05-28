using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML_Tests.Excel.Worksheets
{
    [TestFixture]
    public class XLSheetProtectionTests
    {
        [Test]
        public void TestUnprotectWorksheetWithNoPassword()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\SHA512PasswordProtection.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.IsTrue(ws.Protection.Protected);
                ws.Unprotect();
                Assert.IsFalse(ws.Protection.Protected);
            }
        }

        [Test]
        public void TestWorksheetWithSHA512Protection()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Misc\SHA512PasswordProtection.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet2");
                Assert.IsTrue(ws.Protection.Protected);
                // Protected with SHA-512 password - not yet supported.
                Assert.Throws<ArgumentException>(() => ws.Unprotect());

                // Protected with SHA-512 password - not yet supported.
                Assert.Throws<ArgumentException>(() => ws.Unprotect("abc"));
                Assert.IsTrue(ws.Protection.Protected);
            }
        }
    }
}
