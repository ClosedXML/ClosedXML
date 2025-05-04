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
            using var ms = new MemoryStream();
            using (var stream = GetProtectedWorkbookStreamWithPassword())
            using (var wb = new XLWorkbook(stream))
            {
                Assert.That(wb.Protection.Algorithm, Is.EqualTo(Algorithm.SHA512));
                wb.Unprotect("12345");
                wb.Protect("12345", Algorithm.SimpleHash);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.IsPasswordProtected, Is.True);
                    Assert.That(wb.Protection.Algorithm, Is.EqualTo(Algorithm.SimpleHash));
                });
            }
        }

        [Test]
        public void CanChangeToPasswordProtected()
        {
            using var ms = new MemoryStream();
            using (var stream = GetProtectedWorkbookStreamWithoutPassword())
            using (var wb = new XLWorkbook(stream))

            {
                wb.Unprotect();
                wb.Protection.Protect("12345");

                Assert.That(wb.Protection.IsPasswordProtected, Is.True);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.IsPasswordProtected, Is.True);
                    Assert.That(wb.Protection.Algorithm, Is.EqualTo(Algorithm.SimpleHash));
                });
                Assert.That(wb.Protection.PasswordHash, Is.Not.Empty);
            }
        }

        [Test]
        public void CanChangeToProtectedWithoutPassword()
        {
            using var ms = new MemoryStream();
            using (var stream = GetProtectedWorkbookStreamWithPassword())
            using (var wb = new XLWorkbook(stream))

            {
                wb.Unprotect("12345");
                wb.Protection.Protect();

                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.IsPasswordProtected, Is.False);
                    Assert.That(wb.Protection.IsProtected, Is.True);
                });

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.IsPasswordProtected, Is.False);
                    Assert.That(wb.Protection.IsProtected, Is.True);
                    Assert.That(wb.Protection.Algorithm, Is.EqualTo(Algorithm.SimpleHash));
                    Assert.That(wb.Protection.PasswordHash, Is.Empty);
                });
            }
        }

        [Test]
        public void CannotUnprotectIfNoPassword()
        {
            using var stream = GetProtectedWorkbookStreamWithoutPassword();
            using var wb = new XLWorkbook(stream);
            var ex = Assert.Throws<ArgumentException>(() => wb.Unprotect("dummy password"));
            Assert.That(ex.Message, Is.EqualTo("Invalid password"));
        }

        [Test]
        public void CannotUnprotectWithoutPassword()
        {
            using var stream = GetProtectedWorkbookStreamWithPassword();
            using var wb = new XLWorkbook(stream);
            var ex = Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
            Assert.That(ex.Message, Is.EqualTo("The workbook structure is password protected"));
        }

        [Test]
        [Theory]
        public void CanProtectWithPassword(Algorithm algorithm)
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet();

                Assert.That(wb.Protection.IsProtected, Is.False);

                wb.Protection.Protect("12345", algorithm);

                wb.Protection.AllowNone();
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), Is.False);
                    Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), Is.False);
                });

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.IsPasswordProtected, Is.True);
                    Assert.That(wb.Protection.IsProtected, Is.True);

                    Assert.That(wb.Protection.Algorithm, Is.EqualTo(algorithm));
                });
                Assert.That(wb.Protection.PasswordHash, Is.Not.Empty);

                Assert.Multiple(() =>
                {
                    Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), Is.False);
                    Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), Is.False);
                });

                var ex = Assert.Throws<ArgumentException>(() => wb.Unprotect("dummy password"));
                Assert.That(ex.Message, Is.EqualTo("Invalid password"));

                wb.Protection.Unprotect("12345");

                wb.Save();
            }
        }

        [Test]
        public void CanUnprotectWithoutPassword()
        {
            using var ms = new MemoryStream();
            using (var stream = GetProtectedWorkbookStreamWithoutPassword())
            using (var wb = new XLWorkbook(stream))
            {
                // Unprotect without password
                wb.Unprotect();

                Assert.That(wb.Protection.IsProtected, Is.False);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.Protection.IsProtected, Is.False);
            }
        }

        [Test]
        public void CanUnprotectWithPassword()
        {
            using var ms = new MemoryStream();
            using (var stream = GetProtectedWorkbookStreamWithPassword())
            using (var wb = new XLWorkbook(stream))
            {
                // Unprotect with password
                wb.Unprotect("12345");

                Assert.That(wb.Protection.IsProtected, Is.False);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.Protection.IsProtected, Is.False);
            }
        }

        [Test]
        public void CopyProtectionFromAnotherWorkbook()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\WorkbookProtection.xlsx"));
            using var wb1 = new XLWorkbook(stream);
            using var wb2 = new XLWorkbook();
            wb2.AddWorksheet();

            var p1 = wb1.Protection.CastTo<XLWorkbookProtection>();
            Assert.Multiple(() =>
            {
                Assert.That(p1.IsProtected, Is.True);

                Assert.That(wb2.Protection.IsProtected, Is.False);
            });
            var p2 = wb2.Protection.CopyFrom(wb1.Protection).CastTo<XLWorkbookProtection>();

            Assert.Multiple(() =>
            {
                Assert.That(p2.IsProtected, Is.True);
                Assert.That(p2.IsPasswordProtected, Is.True);
                Assert.That(p2.Algorithm, Is.EqualTo(p1.Algorithm));
                Assert.That(p2.PasswordHash, Is.EqualTo(p1.PasswordHash));
                Assert.That(p2.Base64EncodedSalt, Is.EqualTo(p1.Base64EncodedSalt));
                Assert.That(p2.SpinCount, Is.EqualTo(p1.SpinCount));

                Assert.That(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), Is.True);
                Assert.That(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), Is.False);
            });

            Assert.Throws<InvalidOperationException>(() => wb2.Unprotect());
            wb2.Unprotect("Abc@123");
        }

        [Test]
        public void IXLProtectableTests()
        {
            using var wb = new XLWorkbook();
            Enumerable.Range(1, 5).ForEach(i => wb.AddWorksheet());

            var list = new List<IXLProtectable>() { wb };
            list.AddRange(wb.Worksheets);

            list.ForEach(el => el.Protect());

            list.ForEach(el => Assert.That(el.IsProtected, Is.True));
            list.ForEach(el => Assert.That(el.IsPasswordProtected, Is.False));

            list.ForEach(el => el.Unprotect());

            list.ForEach(el => Assert.That(el.IsProtected, Is.False));
            list.ForEach(el => Assert.That(el.IsPasswordProtected, Is.False));

            list.ForEach(el => el.Protect("password"));

            list.ForEach(el => Assert.That(el.IsProtected, Is.True));
            list.ForEach(el => Assert.That(el.IsPasswordProtected, Is.True));

            list.ForEach(el => el.Unprotect("password"));

            list.ForEach(el => Assert.That(el.IsProtected, Is.False));
            list.ForEach(el => Assert.That(el.IsPasswordProtected, Is.False));
        }

        [Test]
        public void LoadProtectionWithoutPasswordFromFile()
        {
            using var stream = GetProtectedWorkbookStreamWithoutPassword();
            using var wb = new XLWorkbook(stream);
            Assert.Multiple(() =>
            {
                Assert.That(wb.Protection.IsPasswordProtected, Is.False);
                Assert.That(wb.Protection.IsProtected, Is.True);
                Assert.That(wb.Protection.PasswordHash, Is.Empty);
                Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), Is.True);
                Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), Is.False);
            });
        }

        [Test]
        public void LoadProtectionWithPasswordFromFile()
        {
            using var stream = GetProtectedWorkbookStreamWithPassword();
            using var wb = new XLWorkbook(stream);
            Assert.Multiple(() =>
            {
                Assert.That(wb.Protection.IsPasswordProtected, Is.True);
                Assert.That(wb.Protection.PasswordHash, Is.Not.Empty);
            });
            Assert.Multiple(() =>
            {
                Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), Is.True);
                Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), Is.False);
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(wb1.Protection.IsProtected, Is.True);

                Assert.That(wb1.Protection.AllowedElements, Is.EqualTo(XLWorkbookProtectionElements.Windows));
            });

            wb2.Protection = wb1.Protection;

            Assert.Multiple(() =>
            {
                Assert.That(ReferenceEquals(wb1.Protection, wb2.Protection), Is.False);
                Assert.That(wb2.Protection.IsProtected, Is.True);
                Assert.That(wb2.Protection.AllowedElements, Is.EqualTo(XLWorkbookProtectionElements.Windows));
                Assert.That(wb2.Protection.PasswordHash, Is.EqualTo(wb1.Protection.PasswordHash));
            });
        }

        private Stream GetProtectedWorkbookStreamWithoutPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithoutpassword.xlsx"));

        private Stream GetProtectedWorkbookStreamWithPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithpassword.xlsx"));
    }
}
