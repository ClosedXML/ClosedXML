using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class WorkbookProtection : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Workbook Protection");
#pragma warning disable CS0618 // Type or member is obsolete, but still should be tested
            wb.Protect(true, false, "Abc@123");
#pragma warning restore CS0618 // Type or member is obsolete, but still should be tested
            wb.SaveAs(filePath);
        }

        #endregion Methods
    }
}
