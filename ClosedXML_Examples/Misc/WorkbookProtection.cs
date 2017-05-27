using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples.Misc
{
    public class WorkbookProtection : IXLExample
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Workbook Protection");
                wb.Protect(true, false, "Abc@123");
                wb.SaveAs(filePath);
            }
        }

        // Private

        // Override


        #endregion
    }
}
