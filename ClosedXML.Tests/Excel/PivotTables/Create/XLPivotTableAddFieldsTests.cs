using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables.Create
{
    /// <summary>
    /// Tests that add fields to a new empty table. Doesn't test data.
    /// Expected: Make sure Excel can read the stuff we wrote.
    /// </summary>
    [TestFixture]
    internal class XLPivotTableAddFieldsTests
    {
        [Test]
        public void Add_empty_table()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                CreatePivotTableFor2X2(wb);
            }, @"Other\PivotTable\Create\Add_empty_table.xlsx");
        }

        [Test]
        public void Add_one_column_without_value()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var pt = CreatePivotTableFor2X2(wb);

                pt.ColumnLabels.Add("A");
            }, @"Other\PivotTable\Create\Add_one_column_without_value.xlsx");
        }

        [Test]
        public void Add_one_row_without_value()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var pt = CreatePivotTableFor2X2(wb);

                pt.RowLabels.Add("A");
            }, @"Other\PivotTable\Create\Add_one_row_without_value.xlsx");
        }

        [Test]
        public void Add_one_column_and_one_value()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var pt = CreatePivotTableFor2X2(wb);

                pt.ColumnLabels.Add("A");
                pt.Values.Add("B");
            }, @"Other\PivotTable\Create\Add_one_column_and_one_value.xlsx");
        }

        [Test]
        public void Add_one_column_and_two_values()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var pt = CreatePivotTableFor2X2(wb);

                pt.ColumnLabels.Add("A");
                pt.Values.Add("B", "Sum of B").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("B", "Count of B").SetSummaryFormula(XLPivotSummary.Count);
                pt.SetShowGrandTotalsColumns(false);
            }, @"Other\PivotTable\Create\Add_one_column_and_two_values.xlsx");
        }

        private static IXLPivotTable CreatePivotTableFor2X2(XLWorkbook wb)
        {
            var range = wb.AddWorksheet().FirstCell().InsertData(new object[]
            {
                ("A", "B"),
                (1, 2),
            });

            var ws = wb.AddWorksheet().SetTabActive();
            var pt = ws.PivotTables.Add("Test", ws.FirstCell(), range);
            return pt;
        }
    }
}
