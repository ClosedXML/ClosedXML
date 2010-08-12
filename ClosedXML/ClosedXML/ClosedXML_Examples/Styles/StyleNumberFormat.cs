using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ClosedXML_Examples.Styles
{
    public class StyleNumberFormat
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

        #region Constructors

        // Public
        public StyleNumberFormat()
        {

        }


        // Private


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
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style NumberFormat");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "123456.789";
            ws.Cell(ro, co).Style.NumberFormat.Format = "$ #,##0.00";

            ws.Cell(++ro, co).Value = "12.345";
            ws.Cell(ro, co).Style.NumberFormat.Format = "$ #,##0.00";

            ws.Cell(++ro, co).Value = "12.345";
            ws.Cell(ro, co).Style.NumberFormat.NumberFormatId = 3;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
