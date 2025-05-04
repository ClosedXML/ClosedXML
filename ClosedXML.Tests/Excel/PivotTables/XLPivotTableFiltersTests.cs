using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables
{
    [TestFixture]
    public class XLPivotTableFiltersTests
    {
        [Test]
        public void Adding_and_removing_filters_shifts_pivot_table_area()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var data = ws.Cell("A1").InsertData(new object[]
            {
                ("Name", "City", "Flavor", "Value"),
                ("Cake", "Tokyo", "Vanilla", 7),
            });

            var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), data);

            // No filter, the table is at the original cell
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo("E2"));

            pt.ReportFilters.Add("City");

            // First filter also adds divider row between filter and the table.
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo("E4"));

            pt.ReportFilters.Add("Flavor");

            // When second filter is added, there is no need to add second divider row.
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo("E5"));

            pt.ReportFilters.Remove("City");
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo("E4"));

            pt.ReportFilters.Remove("Flavor");
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo("E2"));
        }
    }
}
