using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.Misc
{
    public class DataValidation
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
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data Validation");

            // Decimal between 1 and 5
            ws.Cell(1, 1).DataValidation.Decimal.Between(1, 5);

            // Whole number equals 2
            var dv1 = ws.Range("A2:A3").DataValidation;
            dv1.WholeNumber.EqualTo(2);
            // Change the error message
            dv1.ErrorStyle = XLErrorStyle.Warning;
            dv1.ErrorTitle = "Number out of range";
            dv1.ErrorMessage = "This cell only allows the number 2.";

            // Date after the millenium
            var dv2 = ws.Cell("A4").DataValidation;
            dv2.Date.EqualOrGreaterThan(new DateTime(2000, 1, 1));
            // Change the input message
            dv2.InputTitle = "Can't party like it's 1999.";
            dv2.InputMessage = "Please enter a date in this century.";

            // From a list
            ws.Cell("C1").Value = "Yes";
            ws.Cell("C2").Value = "No";
            ws.Cell("A5").DataValidation.List(ws.Range("C1:C2"));

            // Intersecting dataValidations
            ws.Range("B1:B4").DataValidation.WholeNumber.EqualTo(1);
            ws.Range("B3:B4").DataValidation.WholeNumber.EqualTo(2);

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
