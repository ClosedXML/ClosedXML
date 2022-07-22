using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class PrecedentCellsTests
    {
        [Test]
        public void GetPrecedentRangesPreventsDuplication()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                    var sheet2 = wb.AddWorksheet("Sheet2");
                    var formula = "=MAX(A2:E2)/COUNTBLANK(A2:E2)*MAX(B1:C3)+SUM(Sheet2!B1:C3)+SUM($A$2:$E$2)+A2+B$2+$C$2";

                    var ranges = sheet1.CalcEngine.GetPrecedentRanges(sheet1, formula).ToList();

                    Assert.AreEqual(6, ranges.Count);
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "A2:E2"));
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "B1:C3"));
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet2" && r.ToString() == "B1:C3"));
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "A2:A2"));
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "B$2:B$2"));
                    Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "$C$2:$C$2"));
                }
            }
        }

        [Test]
        public void GetPrecedentRangesDealsWithNamedRanges()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                    sheet1.NamedRanges.Add("NAMED_RANGE", sheet1.Range("A2:B3"));
                    var formula = "=SUM(NAMED_RANGE)";

                    var ranges = sheet1.CalcEngine.GetPrecedentRanges(sheet1, formula).ToList();

                    Assert.AreEqual(1, ranges.Count);
                    Assert.AreEqual("$A$2:$B$3", ranges.Single().ToString());
                }
            }
        }

        [Test]
        public void GetPrecedentCells()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                    var sheet2 = wb.AddWorksheet("Sheet2");
                    var formula = "=MAX(A2:E2)/COUNTBLANK(A2:E2)*MAX(B1:C3)+SUM(Sheet2!B1:C3)+SUM($A$2:$E$2)+A2+B$2+$C$2";
                    var expectedAtSheet1 = new string[]
                        { "A2", "B2", "C2", "D2", "E2", "B1", "C1", "B3", "C3" };
                    var expectedAtSheet2 = new string[]
                        { "B1", "C1", "B2", "C2", "B3", "C3" };

                    var cells = sheet1.CalcEngine.GetPrecedentCells(sheet1, formula).ToList();

                    Assert.AreEqual(15, cells.Count());
                    foreach (var address in expectedAtSheet1)
                    {
                        Assert.IsTrue(cells.Any(cell => cell.Address.Worksheet.Name == sheet1.Name && cell.Address.ToString() == address),
                            string.Format("Address {0}!{1} is not presented", sheet1.Name, address));
                    }
                    foreach (var address in expectedAtSheet2)
                    {
                        Assert.IsTrue(cells.Any(cell => cell.Address.Worksheet.Name == sheet2.Name && cell.Address.ToString() == address),
                            string.Format("Address {0}!{1} is not presented", sheet2.Name, address));
                    }
                }
            }
        }

        [Test]
        public void CanParseWorksheetNamesWithExclamationMark()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet() as XLWorksheet;
                var ws2 = wb.AddWorksheet("Worksheet!");
                var expectedCell = ws2.Cell("B2");

                var cells = ws1.CalcEngine.GetPrecedentCells(ws1, "='Worksheet!'!B2*2");
                Assert.AreSame(expectedCell, cells.Single());
            }
        }
    }
}
