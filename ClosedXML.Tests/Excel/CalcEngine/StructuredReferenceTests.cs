#nullable enable

using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Test cases per <em>[MS-OI29500] 3.2.3.1.1 Structure References</em>.
    /// </summary>
    [TestFixture]
    internal class StructuredReferenceTests
    {
        private static IEnumerable<object[]> TestCases
        {
            get
            {
                // `table-name[]` refers to all cells in table-name except Header Row and Total Row.
                // `table-name[#Data]` refers to all table-name’s cells except Header Row and Total Row. It is equivalent to the form table-name[].
                yield return new object[] { "TableName[]", "E8:H10", "E8:H10" };
                yield return new object[] { "TableName[#Data]", "E8:H10", "E8:H10" };
                yield return new object[] { "tableName[]", "E8:H10", "E8:H10" };

                // table-name[#Headers] refers to all cells in table-name’s Header Row.
                yield return new object[] { "TableName[#Headers]", "E7:H7", "E7:H7" };

                // `table-name[#Total Row] refers to all cells in the table-name’s Total Row
                // No totals -> no area -> #REF!
                yield return new object[] { "TableName[#Totals]", "E11:H11", "#REF!" };

                // `table-name[#All]` refers to the entire table area. table-name[#All] is the union of
                // table-name[#Headers], table-name[#Data], and table-name[#Total Row]
                yield return new object[] { "TableName[#All]", "E7:H11", "E7:H10" };

                // table-name[column-name] refers to all cells in the column named column-name except
                // the cells from Header Row and Total Row.
                // table-name[[column-name]] refers to all cells in the column named column-name except
                // the cells from Header Row and Total Row.
                // table-name[[#Data],[column-name]] is equivalent to table-name[column-name]
                yield return new object[] { "TableName[Second]", "F8:F10", "F8:F10" };
                yield return new object[] { "TableName[second]", "F8:F10", "F8:F10" };
                yield return new object[] { "TableName[[Second]]", "F8:F10", "F8:F10" };
                yield return new object[] { "TableName[[#Data],[Second]]", "F8:F10", "F8:F10" };

                // table-name[[column-name1]:[column-name2]] refers to all cells from column named column-name1
                // through column named column-name2 except the cells from Header Row and Total Row.
                yield return new object[] { "TableName[[Second]:[Fourth]]", "F8:H10", "F8:H10" };
                yield return new object[] { "TableName[[Fourth]:[Second]]", "F8:H10", "F8:H10" };
                yield return new object[] { "tableName[[second]:[fourth]]", "F8:H10", "F8:H10" };

                // table-name[[keyword],[column-name]], where keyword is one of #Headers, #Total Row, #Data, #All,
                // refers to the intersection of the area defined by table-name[keyword] and all cells from the column
                // named column-name.
                yield return new object[] { "TableName[[#Headers],[Second]]", "F7:F7", "F7:F7" };
                yield return new object[] { "TableName[[#Totals],[Second]]", "F11:F11", "#REF!" };
                yield return new object[] { "TableName[[#Data],[Second]]", "F8:F10", "F8:F10" };
                yield return new object[] { "TableName[[#All],[Second]]", "F7:F11", "F7:F10" };

                // table-name[[keyword],[column-name1]:[column-name2]], where keyword is one of #Headers, #Total
                // Row, #Data, #All, refers to the intersection of the area defined by table-name[keyword] and all cells
                // from the table from column named column - name1 through column named column-name2.
                yield return new object[] { "TableName[[#Headers],[Second]:[Fourth]]", "F7:H7", "F7:H7" };
                yield return new object[] { "TableName[[#Totals],[Second]:[Fourth]]", "F11:H11", "#REF!" };
                yield return new object[] { "TableName[[#Data],[Second]:[Fourth]]", "F8:H10", "F8:H10" };
                yield return new object[] { "TableName[[#All],[Second]:[Fourth]]", "F7:H11", "F7:H10" };

                // table-name[[#Headers],[#Data],[column-name]] is the union of table-name[[#Headers],[column-name]]
                // and table-name[[#Data],[column-name]]
                yield return new object[] { "TableName[[#Headers],[#Data],[Third]]", "G7:G10", "G7:G10" };

                // table-name[[#Headers],[#Data],[column-name]] is the union of table-name[[#Headers],[column-name]]
                // and table-name[[#Data],[column-name]]
                yield return new object[] { "TableName[[#Data],[#Totals],[Third]]", "G8:G11", "G8:G10" };

                // table-name[[#Headers],[#Data],[column-name1]:[column-name2]] is the union of
                // table-name[[#Headers], [column-name1]:[column-name2]] and table-name[[#Data],
                // [column-name1]:[column - name2]]
                yield return new object[] { "TableName[[#Headers],[#Data],[Third]:[Fourth]]", "G7:H10", "G7:H10" };
                yield return new object[] { "TableName[[#Headers],[#Data],[Fourth]:[Third]]", "G7:H10", "G7:H10" };

                // table-name[[#Data],[#Total Row], [column-name1]:[column-name2]] is the union of
                // table-name[[#Data], [column-name1]:[column-name2]] and table-name[[#Total Row],
                // [column-name1]:[column-name2]]
                yield return new object[] { "TableName[[#Data],[#Totals],[Second]:[Third]]", "F8:G11", "F8:G10" };
                yield return new object[] { "TableName[[#Data],[#Totals],[Third]:[Second]]", "F8:G11", "F8:G10" };

                // Incorrect name of table or column returns #REF!
                yield return new object[] { "WrongName[]", "#REF!", "#REF!" };
                yield return new object[] { "TableName[[NonExistentCol]]", "#REF!", "#REF!" };
                yield return new object[] { "TableName[[First]:[NonExistentCol]]", "#REF!", "#REF!" };
                yield return new object[] { "TableName[[NonExistentCol]:[Fourth]]", "#REF!", "#REF!" };
                yield return new object[] { "TableName[[NonExistent1]:[NonExistent2]]", "#REF!", "#REF!" };
            }
        }

        [TestCaseSource(nameof(TestCases))]
        public void Structured_reference_is_resolved_to_reference(
            string structuredReference,
            string expectedWithTotals,
            string expectedWithoutTotals)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var table = Add4X3Table(ws, "E7");
            table.ShowTotalsRow = true;

            AssertRange(structuredReference, expectedWithTotals, ws);

            table.ShowTotalsRow = false;
            AssertRange(structuredReference, expectedWithoutTotals, ws);
        }

        [Test]
        public void This_row_of_column_of_table_reference()
        {
            // table-name[[#This Row],[column-name]] refers to the cell in the intersection of table-name[column-
            // name] and the current row; for example, the row of the cell that contains the formula with the
            // structure reference. table-name[[#This Row],[column-name1]:[column-name2]]refers to the cells in
            // the intersection of table-name[[column - name]:[column - name2]] and the current row; for example,
            // the row of the cell that contains the formula with such structure reference.These two forms allow
            //formulas to perform implicit intersection using structure references.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            Add4X3Table(ws, "E7");

            const string columnFormula = "TableName[[#This Row],[Second]]";
            AssertRange(columnFormula, "F8:F8", ws, "D8");
            AssertRange(columnFormula, "F10:F10", ws, "D10");

            const string columnsFormula = "TableName[[#This Row],[Second]:[Third]]";
            AssertRange(columnsFormula, "F8:G8", ws, "D8");
            AssertRange(columnsFormula, "F10:G10", ws, "D10");
        }

        [TestCase("TableName[[#This Row],[Second]]")]
        [TestCase("TableName[[#This Row],[Second]:[Fourth]]")]
        [TestCase("TableName[[#This Row],[Fourth]:[Second]]")]
        public void This_row_outside_data_area_of_table_reference(string formula)
        {
            // table-name[[#This Row],[column-name]] and table-name[[#This Row],[column-name1]:[column-name2]]
            // return #VALUE! when the row is not in data range of rows.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var table = Add4X3Table(ws, "E7");
            table.ShowTotalsRow = true;

            // Right above header row
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula, "D6"));

            // Header row
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula, "D7"));

            // Whether there is a totals row or not, the result is #VALUE!
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula, "D11"));

            table.ShowTotalsRow = false;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate(formula, "D11"));
        }

        private static IXLTable Add4X3Table(IXLWorksheet ws, string origin)
        {
            var dt = new DataTable("TableName");
            dt.Columns.AddRange(new[]
            {
                new DataColumn("First", typeof(int)),
                new DataColumn("Second", typeof(int)),
                new DataColumn("Third", typeof(int)),
                new DataColumn("Fourth", typeof(int)),
            });

            for (var i = 1; i <= 3; ++i)
            {
                var row = dt.NewRow();
                row["First"] = i;
                row["Second"] = i * 10;
                row["Third"] = i * 100;
                row["Fourth"] = i * 1000;
                dt.Rows.Add(row);
            }

            var table = ws.Cell(origin).InsertTable(dt, "TableName");
            table.SetShowTotalsRow(true);
            return table;
        }

        private static void AssertRange(string structureReference, string expectedArea, IXLWorksheet ws, string? formulaAddress = null)
        {
            if (expectedArea == "#REF!")
            {
                Assert.AreEqual(XLError.CellReference, ws.Evaluate(structureReference, formulaAddress));
                return;
            }

            var expected = XLSheetRange.Parse(expectedArea);
            Assert.AreEqual(expected.LeftColumn, ws.Evaluate($"COLUMN({structureReference})", formulaAddress));
            Assert.AreEqual(expected.TopRow, ws.Evaluate($"ROW({structureReference})", formulaAddress));
            Assert.AreEqual(expected.Height, ws.Evaluate($"ROWS({structureReference})", formulaAddress));
            Assert.AreEqual(expected.Width, ws.Evaluate($"COLUMNS({structureReference})", formulaAddress));
        }
    }
}
