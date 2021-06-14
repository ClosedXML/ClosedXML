// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests
{
    public class SortTests
    {
        [Test]
        public void SortIsFast()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 1).Value = "A";
            ws.Cell(1, 2).Value = "B";
            for (int i = 0; i < 14000; i++)
            {
                ws.Cell(i + 2, 1).Value = i;
                ws.Cell(i + 2, 2).Value = i % 2;
            }

            var autoFilter = ws.Range(1, 1, 14001, 2).SetAutoFilter();

            var stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();

            // Before the fix, sorting this used to take about 12min on my laptop
            autoFilter.Sort(2);

            stopwatch.Stop();

            Assert.True(stopwatch.ElapsedMilliseconds < 10000);
        }

        [Test]
        public void SortPreservesFixedFormula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            AddTestTable(ws);

            Assert.AreEqual("LEN($A$1)", ws.Cell("C1").FormulaA1);
            Assert.AreEqual("LEN($A$7)", ws.Cell("C7").FormulaA1);

            ws.RangeUsed().Sort(2, XLSortOrder.Ascending, matchCase: false, ignoreBlanks: true);

            Assert.AreEqual("LEN($A$3)", ws.Cell("C7").FormulaA1);
            Assert.AreEqual("LEN($A$7)", ws.Cell("C1").FormulaA1);
        }

        [Test]
        public void SortPreservesRelativeFormula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            AddTestTable(ws);

            Assert.AreEqual("LEN(A1)", ws.Cell("D1").FormulaA1);
            Assert.AreEqual("LEN(A7)", ws.Cell("D7").FormulaA1);

            ws.RangeUsed().Sort(2, XLSortOrder.Ascending, matchCase: false, ignoreBlanks: true);

            Assert.AreEqual("LEN(A1)", ws.Cell("D1").FormulaA1);
            Assert.AreEqual("LEN(A7)", ws.Cell("D7").FormulaA1);
        }

        private void AddTestTable(IXLWorksheet ws)
        {
            var data = new[] {
                ("B", 5, XLColor.LightGreen),
                ("A", 2, XLColor.DarkTurquoise),
                ("a", 7, XLColor.BurlyWood),
                ("A", 3, XLColor.DarkGray),
                ("", 8, XLColor.DarkSalmon),
                ("A", 4, XLColor.DodgerBlue),
                ("a", 1, XLColor.IndianRed),
                ("B", 6, XLColor.DeepPink)
            };

            Enumerable.Range(1, 8).ForEach(i =>
            {
                var (a, b, color) = data[i - 1];
                ws.Cell(i, 1).SetValue(a).Style.Fill.SetBackgroundColor(color);
                ws.Cell(i, 2).SetValue(b).Style.Fill.SetBackgroundColor(color);
                ws.Cell(i, 3).SetFormulaA1($"LEN($A${i})").Style.Fill.SetBackgroundColor(color);
                ws.Cell(i, 4).SetFormulaA1($"LEN(A{i})").Style.Fill.SetBackgroundColor(color);
            });
        }
    }
}
