// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Tests
{
    public class XLWorkbookProtectionTests
    {
        [Test]
        public void CanChangeProtectionAlgorithm()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = GetProtectedWorkbookStreamWithPassword())
                using (var wb = new XLWorkbook(stream))
                {
                    Assert.AreEqual(Algorithm.SHA512, wb.Protection.Algorithm);
                    wb.Unprotect("12345");
                    wb.Protect("12345", Algorithm.SimpleHash);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsTrue(wb.IsPasswordProtected);
                    Assert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
                }
            }
        }

        [Test]
        public void CanChangeToPasswordProtected()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = GetProtectedWorkbookStreamWithoutPassword())
                using (var wb = new XLWorkbook(stream))

                {
                    wb.Unprotect();
                    wb.Protection.Protect("12345");

                    Assert.IsTrue(wb.Protection.IsPasswordProtected);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsTrue(wb.Protection.IsPasswordProtected);
                    Assert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
                    Assert.AreNotEqual("", wb.Protection.PasswordHash);
                }
            }
        }

        [Test]
        public void CanChangeToProtectedWithoutPassword()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = GetProtectedWorkbookStreamWithPassword())
                using (var wb = new XLWorkbook(stream))

                {
                    wb.Unprotect("12345");
                    wb.Protection.Protect();

                    Assert.IsFalse(wb.Protection.IsPasswordProtected);
                    Assert.IsTrue(wb.Protection.IsProtected);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsFalse(wb.Protection.IsPasswordProtected);
                    Assert.IsTrue(wb.Protection.IsProtected);
                    Assert.AreEqual(Algorithm.SimpleHash, wb.Protection.Algorithm);
                    Assert.AreEqual("", wb.Protection.PasswordHash);
                }
            }
        }

        [Test]
        public void CannotUnprotectIfNoPassword()
        {
            using (var stream = GetProtectedWorkbookStreamWithoutPassword())
            using (var wb = new XLWorkbook(stream))
            {
                var ex = Assert.Throws<ArgumentException>(() => wb.Unprotect("dummy password"));
                Assert.AreEqual("Invalid password", ex.Message);
            }
        }

        [Test]
        public void CannotUnprotectWithoutPassword()
        {
            using (var stream = GetProtectedWorkbookStreamWithPassword())
            using (var wb = new XLWorkbook(stream))
            {
                var ex = Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
                Assert.AreEqual("The workbook structure is password protected", ex.Message);
            }
        }

        [Test]
        [Theory]
        public void CanProtectWithPassword(Algorithm algorithm)
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet();

                    Assert.IsFalse(wb.Protection.IsProtected);

                    wb.Protection.Protect("12345", algorithm);

                    wb.Protection.AllowNone();
                    Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure));
                    Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsTrue(wb.Protection.IsPasswordProtected);
                    Assert.IsTrue(wb.Protection.IsProtected);

                    Assert.AreEqual(algorithm, wb.Protection.Algorithm);
                    Assert.AreNotEqual("", wb.Protection.PasswordHash);

                    Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure));
                    Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));

                    var ex = Assert.Throws<ArgumentException>(() => wb.Unprotect("dummy password"));
                    Assert.AreEqual("Invalid password", ex.Message);

                    wb.Protection.Unprotect("12345");

                    wb.Save();
                }
            }
        }

        [Test]
        public void CanUnprotectWithoutPassword()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = GetProtectedWorkbookStreamWithoutPassword())
                using (var wb = new XLWorkbook(stream))
                {
                    // Unprotect without password
                    wb.Unprotect();

                    Assert.IsFalse(wb.Protection.IsProtected);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsFalse(wb.Protection.IsProtected);
                }
            }
        }

        [Test]
        public void CanUnprotectWithPassword()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = GetProtectedWorkbookStreamWithPassword())
                using (var wb = new XLWorkbook(stream))
                {
                    // Unprotect with password
                    wb.Unprotect("12345");

                    Assert.IsFalse(wb.Protection.IsProtected);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.IsFalse(wb.Protection.IsProtected);
                }
            }
        }

        [Test]
        public void CopyProtectionFromAnotherWorkbook()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\WorkbookProtection.xlsx")))
            using (var wb1 = new XLWorkbook(stream))
            using (var wb2 = new XLWorkbook())
            {
                wb2.AddWorksheet();

                var p1 = wb1.Protection.CastTo<XLWorkbookProtection>();
                Assert.IsTrue(p1.IsProtected);

                Assert.IsFalse(wb2.Protection.IsProtected);
                var p2 = wb2.Protection.CopyFrom(wb1.Protection).CastTo<XLWorkbookProtection>();

                Assert.IsTrue(p2.IsProtected);
                Assert.IsTrue(p2.IsPasswordProtected);
                Assert.AreEqual(p1.Algorithm, p2.Algorithm);
                Assert.AreEqual(p1.PasswordHash, p2.PasswordHash);
                Assert.AreEqual(p1.Base64EncodedSalt, p2.Base64EncodedSalt);
                Assert.AreEqual(p1.SpinCount, p2.SpinCount);

                Assert.IsTrue(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));
                Assert.IsFalse(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure));

                Assert.Throws<InvalidOperationException>(() => wb2.Unprotect());
                wb2.Unprotect("Abc@123");
            }
        }

        [Test]
        public void IXLProtectableTests()
        {
            using var wb = new XLWorkbook();
            Enumerable.Range(1, 5).ForEach(i => wb.AddWorksheet());

            var list = new List<IXLProtectable>() { wb };
            list.AddRange(wb.Worksheets);

            list.ForEach(el => el.Protect());

            list.ForEach(el => Assert.IsTrue(el.IsProtected));
            list.ForEach(el => Assert.IsFalse(el.IsPasswordProtected));

            list.ForEach(el => el.Unprotect());

            list.ForEach(el => Assert.IsFalse(el.IsProtected));
            list.ForEach(el => Assert.IsFalse(el.IsPasswordProtected));

            list.ForEach(el => el.Protect("password"));

            list.ForEach(el => Assert.IsTrue(el.IsProtected));
            list.ForEach(el => Assert.IsTrue(el.IsPasswordProtected));

            list.ForEach(el => el.Unprotect("password"));

            list.ForEach(el => Assert.IsFalse(el.IsProtected));
            list.ForEach(el => Assert.IsFalse(el.IsPasswordProtected));
        }

        [Test]
        public void LoadProtectionWithoutPasswordFromFile()
        {
            using (var stream = GetProtectedWorkbookStreamWithoutPassword())
            using (var wb = new XLWorkbook(stream))
            {
                Assert.IsFalse(wb.Protection.IsPasswordProtected);
                Assert.IsTrue(wb.Protection.IsProtected);
                Assert.AreEqual("", wb.Protection.PasswordHash);
                Assert.IsTrue(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));
                Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure));
            }
        }

        [Test]
        public void LoadProtectionWithPasswordFromFile()
        {
            using (var stream = GetProtectedWorkbookStreamWithPassword())
            using (var wb = new XLWorkbook(stream))
            {
                Assert.IsTrue(wb.Protection.IsPasswordProtected);
                Assert.AreNotEqual("", wb.Protection.PasswordHash);
                Assert.IsTrue(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows));
                Assert.IsFalse(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure));
            }
        }

        [Test]
        public void SetWorkbookProtectionCloning()
        {
            var wb1 = new XLWorkbook();
            var wb2 = new XLWorkbook();

            wb1.AddWorksheet();
            wb2.AddWorksheet();

            wb1.Protect("123", Algorithm.SHA512)
                .AllowElement(XLWorkbookProtectionElements.Windows)
                .DisallowElement(XLWorkbookProtectionElements.Structure);

            Assert.IsTrue(wb1.Protection.IsProtected);

            Assert.AreEqual(XLWorkbookProtectionElements.Windows, wb1.Protection.AllowedElements);

            wb2.Protection = wb1.Protection;

            Assert.IsFalse(ReferenceEquals(wb1.Protection, wb2.Protection));
            Assert.IsTrue(wb2.Protection.IsProtected);
            Assert.AreEqual(XLWorkbookProtectionElements.Windows, wb2.Protection.AllowedElements);
            Assert.AreEqual(wb1.Protection.PasswordHash, wb2.Protection.PasswordHash);
        }

        private Stream GetProtectedWorkbookStreamWithoutPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithoutpassword.xlsx"));

        private Stream GetProtectedWorkbookStreamWithPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithpassword.xlsx"));
    }
}
