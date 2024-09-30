using ClosedXML.Excel;
using NUnit.Framework;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class PrintAreaTests
    {
        [Test]
        [TestCase("Sheet1!$A$1:$B$2", "A1:B2")]
        [TestCase("Sheet1!$A$1:$B$2,Sheet1!$D$3:$D$5", "A1:B2", "D3:D5")]
        public void CanLoadWorksheetWithMultiplePrintAreas(string expectedRange, params string[] printAreaRangeAddresses)
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    foreach (var printAreaRangeAddress in printAreaRangeAddresses)
                        ws.PageSetup.PrintAreas.Add(printAreaRangeAddress);
                },
                (_, ws) =>
                {
                    Assert.AreEqual(expectedRange, ws.PageSetup.PrintAreas.PrintArea);
                });
        }

        [Test]
        [TestCase("Sheet1!$A$1:$B$2")]
        [TestCase("Sheet1!$A$1:$B$2", "Sheet2!$A$1:$B$2,Sheet2!$D$3:$D$5")]
        public void CanLoadWorksheetWithMultiplePrintAreas_Formulas(params string[] formulas)
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    foreach (var formula in formulas)
                        ws.PageSetup.PrintAreas.AddFormula(formula);
                },
                (_, ws) =>
                {
                    Assert.AreEqual(string.Join(",", formulas), ws.PageSetup.PrintAreas.PrintArea);
                });
        }

        [Test]
        public void RenameWorkSheet()
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    ws.PageSetup.PrintAreas.Add("A1:B2");
                    ws.Name = "NewSheetName";
                },
                (_, ws) =>
                {
                    Assert.AreEqual("NewSheetName!$A$1:$B$2", ws.PageSetup.PrintAreas.PrintArea);
                });
        }

        [Test]
        public void CopyWorksheet()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var sheet1 = wb.AddWorksheet("Sheet1");
                sheet1.PageSetup.PrintAreas.Add("A1:B2");

                sheet1.CopyTo("Sheet2");

                wb.SaveAs(ms, true, true);
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.TryGetWorksheet("Sheet1", out XLWorksheet sheet1));
                Assert.That(wb.TryGetWorksheet("Sheet2", out XLWorksheet sheet2));

                Assert.AreEqual("Sheet1!$A$1:$B$2", sheet1.PageSetup.PrintAreas.PrintArea);
                Assert.AreEqual("Sheet2!$A$1:$B$2", sheet2.PageSetup.PrintAreas.PrintArea);
            }
        }
    }
}
