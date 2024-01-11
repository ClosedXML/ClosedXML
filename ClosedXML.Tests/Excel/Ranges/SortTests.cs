using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges
{
    [TestFixture]
    public class SortTests
    {
        [Test]
        public void Values_are_sorted_by_type_first()
        {
            // The values in asc order are number, text, logical, error, blanks.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var values = new XLCellValue[]
            {
                1,
                "",
                "#VALUE!",
                "1",
                "Text",
                "TRUE",
                true,
                XLError.IncompatibleValue,
                Blank.Value,
            };

            // Assign in reverse order
            for (var row = 1; row <= values.Length; ++row)
                ws.Cell(row, 1).Value = values[^row];

            ws.Range(1, 1, values.Length, 1).Sort("1 ASC");

            for (var row = 1; row <= values.Length; ++row)
            {
                var sortedValue = ws.Cell(row, 1).Value;
                Assert.AreEqual(values[row - 1], sortedValue);
            }
        }

        [TestCase(XLSortOrder.Ascending)]
        [TestCase(XLSortOrder.Descending)]
        public void Blanks_are_always_last(XLSortOrder sortOrder)
        {
            // When range contains blank, it is always last, no matter
            // if the sort order is ascending or descending
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var values = new XLCellValue[]
            {
                1,
                Blank.Value,
                2,
            };
            for (var row = 1; row <= values.Length; ++row)
                ws.Cell(row, 1).Value = values[row - 1];

            ws.Range(1, 1, values.Length, 1).Sort("1", sortOrder);

            Assert.AreEqual(Blank.Value, ws.Cell(3, 1).Value);
        }

        [Test]
        public void IgnoreBlanks_set_to_false_treats_blanks_as_empty_strings()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = "Text";
            ws.Cell("A2").Value = Blank.Value;
            ws.Cell("A3").Value = string.Empty;

            ws.Range("A1:A3").Sort(1, ignoreBlanks: false);

            // Since blank is treated as empty string, it is not shuffled to the end.
            Assert.AreEqual(Blank.Value, ws.Cell("A1").Value);
            Assert.AreEqual(string.Empty, ws.Cell("A2").Value);
            Assert.AreEqual("Text", ws.Cell("A3").Value);
        }

        [TestCase(true, "a", "A")]
        [TestCase(false, "A", "a")]
        [Culture("en-US")]
        public void MatchCase_flag_determines_if_texts_are_compared_case_sensitive(bool matchCase, string expectedFirst, string expectedSecond)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // In US locale, lower-case is before upper case.
            ws.Cell("A1").Value = "A";
            ws.Cell("A2").Value = "a";

            ws.Range("A1:A2").Sort(1, matchCase: matchCase);

            Assert.AreEqual(expectedFirst, ws.Cell("A1").Value);
            Assert.AreEqual(expectedSecond, ws.Cell("A2").Value);
        }

        [Test]
        public void Sort_can_use_multiple_columns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertData(new object[]
            {
                new [] { 1, 2 },
                new [] { 2, 2 },
                new [] { 1, 1 },
            });

            ws.Range("A1:B4").Sort("2 ASC, 1 DESC");

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(1, ws.Cell("B1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.AreEqual(2, ws.Cell("B2").Value);
            Assert.AreEqual(1, ws.Cell("A3").Value);
            Assert.AreEqual(2, ws.Cell("B3").Value);
        }

        [Test]
        public void Sort_columns_in_range_by_rows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertData(new object[]
            {
                new [] { 2, 2, 1 },
                new [] { 1, 2, 1 },
            });

            // Doesn't have parameters, so it is first rows ASC, second row ASC.
            ws.Range("A1:C2").SortLeftToRight();

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(1, ws.Cell("A2").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(1, ws.Cell("B2").Value);
            Assert.AreEqual(2, ws.Cell("C1").Value);
            Assert.AreEqual(2, ws.Cell("C2").Value);
        }
    }
}
