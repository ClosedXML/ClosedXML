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
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\CellValues.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            IXLCells foundCells;

            foundCells = ws.Search("Initial Value");
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(1));
                Assert.That(foundCells.Single().Address.ToString(), Is.EqualTo("B2"));
            });
            Assert.That(foundCells.Single().GetText(), Is.EqualTo("Initial Value"));

            foundCells = ws.Search("Using");
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(2));
                Assert.That(foundCells.First().Address.ToString(), Is.EqualTo("D2"));
            });
            Assert.That(foundCells.First().GetText(), Is.EqualTo("Using Get...()"));
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(2));
                Assert.That(foundCells.Last().Address.ToString(), Is.EqualTo("E2"));
            });
            Assert.That(foundCells.Last().GetText(), Is.EqualTo("Using GetValue<T>()"));

            foundCells = ws.Search("1234");
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(5));
                Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()), Is.EqualTo("B5,C5,D5,E5,F5"));
            });

            foundCells = ws.Search("Sep");
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(1));
                Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()), Is.EqualTo("G3"));
            });

            foundCells = ws.Search("1234", CompareOptions.Ordinal, true);
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(5));
                Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()), Is.EqualTo("B5,C5,D5,E5,F5"));
            });

            foundCells = ws.Search("test case", CompareOptions.Ordinal);
            Assert.That(foundCells.Count(), Is.EqualTo(0));

            foundCells = ws.Search("test case", CompareOptions.OrdinalIgnoreCase);
            Assert.That(foundCells.Count(), Is.EqualTo(6));
        }

        [Test]
        public void TestSearch2()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            IXLCells foundCells;

            foundCells = ws.Search("3", CompareOptions.Ordinal);
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(10));
                Assert.That(foundCells.First().Address.ToString(), Is.EqualTo("C2"));
            });

            foundCells = ws.Search("A2", CompareOptions.Ordinal, true);
            Assert.That(foundCells.Count(), Is.EqualTo(6));
            Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()), Is.EqualTo("C2,D2,B6,C6,D6,A11"));

            foundCells = ws.Search("RC", CompareOptions.Ordinal, true);
            Assert.Multiple(() =>
            {
                Assert.That(foundCells.Count(), Is.EqualTo(3));
                Assert.That(string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray()), Is.EqualTo("E2,E3,E4"));
            });
        }
    }
}
