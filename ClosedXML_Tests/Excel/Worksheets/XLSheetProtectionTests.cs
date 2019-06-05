// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel.Worksheets
{
    [TestFixture]
    public class XLSheetProtectionTests
    {
        [Test]
        public void AllowEverything()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowedElements = XLSheetProtectionElements.Everything;

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements)).Cast<XLSheetProtectionElements>())
                    Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowElement(XLSheetProtectionElements.Everything);

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements)).Cast<XLSheetProtectionElements>())
                    Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowEverything();

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements)).Cast<XLSheetProtectionElements>())
                    Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }

        [Test]
        public void AllowNothing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowedElements = XLSheetProtectionElements.None;

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements))
                    .Cast<XLSheetProtectionElements>()
                    .Where(e => e != XLSheetProtectionElements.None))

                    Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowNone();

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements))
                    .Cast<XLSheetProtectionElements>()
                    .Where(e => e != XLSheetProtectionElements.None))

                    Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
            }
        }

        [Test]
        public void CopyProtectionFromAnotherSheet()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\SheetProtection.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws1 = wb.Worksheet("Protected Password = 123");
                Assert.IsTrue(ws1.Protection.IsProtected);

                var ws2 = ws1.CopyTo("New worksheet");
                Assert.IsFalse(ws2.Protection.IsProtected);
                ws2.Protection.CopyFrom(ws1.Protection);
                Assert.IsTrue(ws2.Protection.IsProtected);
                Assert.IsTrue(ws2.Protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns));
                Assert.IsTrue(ws2.Protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows));
                Assert.IsFalse(ws2.Protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks));

                Assert.Throws<ArgumentException>(() => ws2.Unprotect());
                ws2.Unprotect("123");
            }
        }

        [Test]
        public void TestUnprotectWorksheetWithNoPassword()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.IsTrue(ws.Protection.IsProtected);
                ws.Unprotect();
                Assert.IsFalse(ws.Protection.IsProtected);
            }
        }

        [Test]
        public void TestWorksheetWithSHA512Protection()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet2");
                Assert.IsTrue(ws.Protection.IsProtected);
                // Protected with SHA-512 password - not yet supported.
                Assert.Throws<ArgumentException>(() => ws.Unprotect());

                // Protected with SHA-512 password - not yet supported.
                Assert.Throws<ArgumentException>(() => ws.Unprotect("abc"));
                Assert.IsTrue(ws.Protection.IsProtected);
            }
        }

        [Test]
        public void SetWorksheetProtectionCloning()
        {
            var ws1 = new XLWorkbook().AddWorksheet();
            var ws2 = new XLWorkbook().AddWorksheet();

            ws1.Protect("123")
                .AllowElement(XLSheetProtectionElements.FormatEverything)
                .DisallowElement(XLSheetProtectionElements.FormatCells);

            Assert.AreEqual(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything, ws1.Protection.AllowedElements);

            ws2.Protection = ws1.Protection;

            Assert.IsFalse(ReferenceEquals(ws1.Protection, ws2.Protection));
            Assert.IsTrue(ws2.Protection.IsProtected);
            Assert.AreEqual(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything, ws2.Protection.AllowedElements);
            Assert.AreEqual((ws1.Protection as XLSheetProtection).PasswordHash, (ws2.Protection as XLSheetProtection).PasswordHash);
        }
    }
}
