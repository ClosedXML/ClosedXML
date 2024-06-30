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
            Assert.AreEqual("E2", ((XLPivotTable)pt).Area.ToString());

            pt.ReportFilters.Add("City");

            // First filter also adds divider row between filter and the table.
            Assert.AreEqual("E4", ((XLPivotTable)pt).Area.ToString());

            pt.ReportFilters.Add("Flavor");

            // When second filter is added, there is no need to add second divider row.
            Assert.AreEqual("E5", ((XLPivotTable)pt).Area.ToString());

            pt.ReportFilters.Remove("City");
            Assert.AreEqual("E4", ((XLPivotTable)pt).Area.ToString());

            pt.ReportFilters.Remove("Flavor");
            Assert.AreEqual("E2", ((XLPivotTable)pt).Area.ToString());
        }
    }
}
