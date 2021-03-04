using ClosedXML.Excel;
using NUnit.Framework;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class SearchTests
    {
        [Test]
        public void TestSearch()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\CellValues.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();

                IXLCells foundCells;

                foundCells = ws.Search("Initial Value");
                Assert.AreEqual(1, foundCells.Count());
                Assert.AreEqual("B2", foundCells.Single().Address.ToString());
                Assert.AreEqual("Initial Value", foundCells.Single().GetString());

                foundCells = ws.Search("Using");
                Assert.AreEqual(2, foundCells.Count());
                Assert.AreEqual("D2", foundCells.First().Address.ToString());
                Assert.AreEqual("Using Get...()", foundCells.First().GetString());
                Assert.AreEqual(2, foundCells.Count());
                Assert.AreEqual("E2", foundCells.Last().Address.ToString());
                Assert.AreEqual("Using GetValue<T>()", foundCells.Last().GetString());

                foundCells = ws.Search("1234");
                Assert.AreEqual(4, foundCells.Count());
                Assert.AreEqual("C5,D5,E5,F5", string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()));

                foundCells = ws.Search("Sep");
                Assert.AreEqual(2, foundCells.Count());
                Assert.AreEqual("B3,G3", string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()));

                foundCells = ws.Search("1234", CompareOptions.Ordinal, true);
                Assert.AreEqual(5, foundCells.Count());
                Assert.AreEqual("B5,C5,D5,E5,F5", string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()));

                foundCells = ws.Search("test case", CompareOptions.Ordinal);
                Assert.AreEqual(0, foundCells.Count());

                foundCells = ws.Search("test case", CompareOptions.OrdinalIgnoreCase);
                Assert.AreEqual(6, foundCells.Count());
            }
        }

        [Test]
        public void TestSearch2()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();

                IXLCells foundCells;

                foundCells = ws.Search("3", CompareOptions.Ordinal);
                Assert.AreEqual(10, foundCells.Count());
                Assert.AreEqual("C2", foundCells.First().Address.ToString());

                foundCells = ws.Search("A2", CompareOptions.Ordinal, true);
                Assert.AreEqual(6, foundCells.Count());
                Assert.AreEqual("C2,D2,B6,C6,D6,A11", string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()));

                foundCells = ws.Search("RC", CompareOptions.Ordinal, true);
                Assert.AreEqual(3, foundCells.Count());
                Assert.AreEqual("E2,E3,E4", string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()));
            }
        }
    }
}
