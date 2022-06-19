using ClosedXML.Excel;

namespace ClosedXML.Examples.Ranges
{
    public class SelectingRanges : IXLExample
    {
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var wsActiveCell = wb.AddWorksheet("Set Active Cell");
            wsActiveCell.Cell("B2").SetActive();

            var wsSelectRowsColumns = wb.AddWorksheet("Select Rows and Columns");
            wsSelectRowsColumns.Rows("2, 4-5").Select();
            wsSelectRowsColumns.Columns("2, 4-5").Select();

            var wsSelectMisc = wb.AddWorksheet("Select Misc");
            wsSelectMisc.Cell("B2").Select();
            wsSelectMisc.Range("D2:E2").Select();
            wsSelectMisc.Ranges("C3, D4:E5").Select();

            wb.SaveAs(filePath);
        }
    }
}