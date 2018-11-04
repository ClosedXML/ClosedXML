using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples.Misc
{
    public class DataValidation : IXLExample
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
            ws.Cell(1, 1).SetDataValidation().Decimal.Between(1, 5);

            // Whole number equals 2
            var dv1 = ws.Range("A2:A3").SetDataValidation();
            dv1.WholeNumber.EqualTo(2);
            // Change the error message
            dv1.ErrorStyle = XLErrorStyle.Warning;
            dv1.ErrorTitle = "Number out of range";
            dv1.ErrorMessage = "This cell only allows the number 2.";

            // Date after the millenium
            var dv2 = ws.Cell("A4").SetDataValidation();
            dv2.Date.EqualOrGreaterThan(new DateTime(2000, 1, 1));
            // Change the input message
            dv2.InputTitle = "Can't party like it's 1999.";
            dv2.InputMessage = "Please enter a date in this century.";

            // From a list
            ws.Cell("C1").Value = "Yes";
            ws.Cell("C2").Value = "No";
            ws.Cell("A5").SetDataValidation().List(ws.Range("C1:C2"));

            ws.Range("C1:C2").AddToNamed("YesNo");
            ws.Cell("A6").SetDataValidation().List("=YesNo");

            // Intersecting dataValidations
            ws.Range("B1:B4").SetDataValidation().WholeNumber.EqualTo(1);
            ws.Range("B3:B4").SetDataValidation().WholeNumber.EqualTo(2);


            // Validate with multiple ranges
            var ws2 = wb.Worksheets.Add("Validate Ranges");
            var rng1 = ws2.Ranges("A1:B2,B4:D7,F4:G5");
            rng1.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng1Validation = rng1.SetDataValidation();
            rng1Validation.Decimal.EqualTo(1);
            rng1Validation.IgnoreBlanks = false;

            var rng2 = ws2.Range("A11:E14");
            rng2.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng2Validation = rng2.SetDataValidation();
            rng2Validation.Decimal.EqualTo(2);
            rng2Validation.IgnoreBlanks = false;

            var rng3 = ws2.Range("B2:B12");
            //rng3.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng3Validation = rng3.SetDataValidation();
            rng3Validation.Decimal.EqualTo(3);
            rng3Validation.IgnoreBlanks = true;
            
            var rng4 = ws2.Range("D5:D6");
            //rng4.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng4Validation = rng4.SetDataValidation();
            rng4Validation.Decimal.EqualTo(4);
            rng4Validation.IgnoreBlanks = true;

            var rng5 = ws2.Range("C13:C14");
            //rng5.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng5Validation = rng5.SetDataValidation();
            rng5Validation.Decimal.EqualTo(5);
            rng5Validation.IgnoreBlanks = true;

            var rng6 = ws2.Range("D11:D12");
            //rng6.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng6Validation = rng6.SetDataValidation();
            rng6Validation.Decimal.EqualTo(5);
            rng6Validation.IgnoreBlanks = true;

            var rng7 = ws2.Range("G4:G5");
            //rng7.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng7Validation = rng7.SetDataValidation();
            rng7Validation.Decimal.EqualTo(5);
            rng7Validation.IgnoreBlanks = true;

            ws.CopyTo(ws.Name + " - Copy");
            ws2.CopyTo(ws2.Name + " - Copy");

            wb.AddWorksheet("Copy From Range 1").FirstCell().Value = ws.RangeUsed(XLCellsUsedOptions.All);
            wb.AddWorksheet("Copy From Range 2").FirstCell().Value = ws2.RangeUsed(XLCellsUsedOptions.All);

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
