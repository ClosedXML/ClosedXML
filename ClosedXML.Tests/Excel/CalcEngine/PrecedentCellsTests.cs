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
            using (XLWorkbook wb = new XLWorkbook())
            {
                var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                var sheet2 = wb.AddWorksheet("Sheet2");
                var formula = "=MAX(A2:E2)/COUNTBLANK(A2:E2)*MAX(B1:C3)+SUM(Sheet2!B1:C3)+SUM($A$2:$E$2)+A2+B$2+$C$2";

                var ranges = sheet1.CalcEngine.GetPrecedentRanges(formula, sheet1).ToList();

                Assert.AreEqual(6, ranges.Count);
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "A2:E2"));
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "B1:C3"));
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet2" && r.ToString() == "B1:C3"));
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "A2:A2"));
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "B$2:B$2"));
                Assert.IsTrue(ranges.Any(r => r.Worksheet.Name == "Sheet1" && r.ToString() == "$C$2:$C$2"));
            }
        }

        // TODO: Root
        [Test]
        public void GetPrecedentRangesDealsWithNamedRanges()
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                sheet1.NamedRanges.Add("NAMED_RANGE", sheet1.Range("A2:B3"));
                var formula = "=SUM(NAMED_RANGE)";

                var ranges = sheet1.CalcEngine.GetPrecedentRanges(formula, sheet1).ToList();

                Assert.AreEqual(1, ranges.Count);
                Assert.AreEqual("$A$2:$B$3", ranges.Single().ToString());
            }
        }

        [TestCase("=A1", new[] { "A1" }, new string[] { })]
        [TestCase(
            "=MAX(A2:E2)/COUNTBLANK(A2:E2)*MAX(B1:C3)+SUM(Sheet2!B1:C3)+SUM($A$2:$E$2)+A2+B$2+$C$2",
            new[] { "A2", "B2", "C2", "D2", "E2", "B1", "C1", "B3", "C3" },
            new[] { "B1", "C1", "B2", "C2", "B3", "C3" })]
        public void GetPrecedentCells(string formula, string[] expectedAtSheet1, string[] expectedAtSheet2)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var sheet1 = wb.AddWorksheet("Sheet1") as XLWorksheet;
                var sheet2 = wb.AddWorksheet("Sheet2");

                var remotelyReliable = sheet1.CalcEngine.TryGetPrecedentCells(formula, sheet1, out var cells);

                Assert.True(remotelyReliable);
                Assert.AreEqual(expectedAtSheet1.Length + expectedAtSheet2.Length, cells.Count());
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

        [Test]
        public void CanParseWorksheetNamesWithExclamationMark()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet() as XLWorksheet;
                var ws2 = wb.AddWorksheet("Worksheet!");
                var expectedCell = ws2.Cell("B2");

                var remotelyReliable = ws1.CalcEngine.TryGetPrecedentCells("='Worksheet!'!B2*2", ws1, out var cells);
                Assert.True(remotelyReliable);
                Assert.AreSame(expectedCell, cells.Single());
            }
        }

        [Test]
        public void NamedRangesMeanNonreliablePrecedentCells()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var remotelyReliable = ws.CalcEngine.TryGetPrecedentCells("=IF(A1, SomeRange, 1)", ws, out var cells);
            Assert.False(remotelyReliable);

            ws.Range("B1").AddToNamed("ExistingRange");
            remotelyReliable = ws.CalcEngine.TryGetPrecedentCells("=ExistingRange", ws, out cells);
            Assert.False(remotelyReliable);
        }

        [Test]
        public void NonexistentSheetsMeanUnreliablePrecednetCells()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var remotelyReliable = ws.CalcEngine.TryGetPrecedentCells("=Sheet2!A1", ws, out var cells);
            Assert.False(remotelyReliable);
        }
    }
}
