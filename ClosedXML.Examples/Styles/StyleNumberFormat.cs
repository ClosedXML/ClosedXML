using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleNumberFormat : IXLExample
    {
        #region Variables

        // Public

        // Private

        #endregion Variables

        #region Properties

        // Public

        // Private

        // Override

        #endregion Properties

        #region Constructors

        // Public
        public StyleNumberFormat()
        {
        }

        // Private

        #endregion Constructors

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style NumberFormat");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "123456.789";
            ws.Cell(ro, co).Style.NumberFormat.Format = "$ #,##0.00";

            ws.Cell(++ro, co).Value = "12.345";
            ws.Cell(ro, co).Style.NumberFormat.Format = "0000";

            ws.Cell(++ro, co).Value = "12.345";
            ws.Cell(ro, co).Style.NumberFormat.NumberFormatId = 3;

            ws.Column(co).AdjustToContents();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}