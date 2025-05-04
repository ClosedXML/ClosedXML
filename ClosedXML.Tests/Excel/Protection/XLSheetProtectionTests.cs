// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Tests.Excel.Worksheets
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
                    Assert.That(ws.Protection.AllowedElements.HasFlag(element), Is.True, element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowElement(XLSheetProtectionElements.Everything);

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements)).Cast<XLSheetProtectionElements>())
                    Assert.That(ws.Protection.AllowedElements.HasFlag(element), Is.True, element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowEverything();

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements)).Cast<XLSheetProtectionElements>())
                    Assert.That(ws.Protection.AllowedElements.HasFlag(element), Is.True, element.ToString());
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

                    Assert.That(ws.Protection.AllowedElements.HasFlag(element), Is.False, element.ToString());
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Protect().AllowNone();

                foreach (var element in Enum.GetValues(typeof(XLSheetProtectionElements))
                    .Cast<XLSheetProtectionElements>()
                    .Where(e => e != XLSheetProtectionElements.None))

                    Assert.That(ws.Protection.AllowedElements.HasFlag(element), Is.False, element.ToString());
            }
        }

        [Test]
        public void ChangeHashingAlgorithm()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.Protect("123", Algorithm.SimpleHash);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Protection.IsProtected, Is.True);
                    Assert.That(ws.Protection.Algorithm, Is.EqualTo(Algorithm.SimpleHash));
                });

                ws.Unprotect("123");
                ws.Protect("123", Algorithm.SHA512);
                wb.Save();
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Protection.IsProtected, Is.True);
                    Assert.That(ws.Protection.Algorithm, Is.EqualTo(Algorithm.SHA512));
                });

                Assert.DoesNotThrow(() => ws.Unprotect("123"));
            }
        }

        [Test]
        public void CopyProtectionFromAnotherSheet()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\SheetProtection.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws1 = wb.Worksheet("Protected Password = 123");
            var p1 = ws1.Protection.CastTo<XLSheetProtection>();
            Assert.That(p1.IsProtected, Is.True);

            var ws2 = ws1.CopyTo("New worksheet");
            Assert.That(ws2.Protection.IsProtected, Is.False);
            var p2 = ws2.Protection.CopyFrom(p1).CastTo<XLSheetProtection>();

            Assert.Multiple(() =>
            {
                Assert.That(p2.IsProtected, Is.True);
                Assert.That(p2.IsPasswordProtected, Is.True);
                Assert.That(p2.Algorithm, Is.EqualTo(p1.Algorithm));
                Assert.That(p2.PasswordHash, Is.EqualTo(p1.PasswordHash));
                Assert.That(p2.Base64EncodedSalt, Is.EqualTo(p1.Base64EncodedSalt));
                Assert.That(p2.SpinCount, Is.EqualTo(p1.SpinCount));

                Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns), Is.True);
                Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows), Is.True);
                Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks), Is.False);
            });

            Assert.Throws<InvalidOperationException>(() => ws2.Unprotect());
            ws2.Unprotect("123");
        }

        [Test]
        public void SetWorksheetProtectionCloning()
        {
            var ws1 = new XLWorkbook().AddWorksheet();
            var ws2 = new XLWorkbook().AddWorksheet();

            ws1.Protect("123")
                .AllowElement(XLSheetProtectionElements.FormatEverything)
                .DisallowElement(XLSheetProtectionElements.FormatCells);

            Assert.That(ws1.Protection.AllowedElements, Is.EqualTo(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything));

            ws2.Protection = ws1.Protection;

            Assert.Multiple(() =>
            {
                Assert.That(ReferenceEquals(ws1.Protection, ws2.Protection), Is.False);
                Assert.That(ws2.Protection.IsProtected, Is.True);
                Assert.That(ws2.Protection.AllowedElements, Is.EqualTo(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything));
                Assert.That((ws2.Protection as XLSheetProtection).PasswordHash, Is.EqualTo((ws1.Protection as XLSheetProtection).PasswordHash));
            });
        }

        [Test]
        public void TestUnprotectWorksheetWithNoPassword()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("Sheet1");
            Assert.That(ws.Protection.IsProtected, Is.True);
            ws.Unprotect();
            Assert.That(ws.Protection.IsProtected, Is.False);
        }

        [Test]
        public void TestWorksheetWithSHA512Protection()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("Sheet2");
            Assert.That(ws.Protection.IsProtected, Is.True);

            // Password required
            Assert.Throws<InvalidOperationException>(() => ws.Unprotect());

            Assert.That(ws.Protection.Algorithm, Is.EqualTo(Algorithm.SHA512));
            ws.Unprotect("abc");
            Assert.That(ws.Protection.IsProtected, Is.False);
        }
    }
}
