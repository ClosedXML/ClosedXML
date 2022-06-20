using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleAlignment : IXLExample
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
        public StyleAlignment()
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
            var ws = workbook.Worksheets.Add("Style Alignment");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "Horizontal = Right";
            ws.Cell(ro, co).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Cell(++ro, co).Value = "Indent = 2";
            ws.Cell(ro, co).Style.Alignment.Indent = 2;

            ws.Cell(++ro, co).Value = "JustifyLastLine = true";
            ws.Cell(ro, co).Style.Alignment.JustifyLastLine = true;

            ws.Cell(++ro, co).Value = "ReadingOrder = ContextDependent";
            ws.Cell(ro, co).Style.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent;

            ws.Cell(++ro, co).Value = "RelativeIndent = 2";
            ws.Cell(ro, co).Style.Alignment.RelativeIndent = 2;

            ws.Cell(++ro, co).Value = "ShrinkToFit = true";
            ws.Cell(ro, co).Style.Alignment.ShrinkToFit = true;

            ws.Cell(++ro, co).Value = "TextRotation = 45";
            ws.Cell(ro, co).Style.Alignment.TextRotation = 45;

            ws.Cell(++ro, co).Value = "TopToBottom = true";
            ws.Cell(ro, co).Style.Alignment.TopToBottom = true;

            ws.Cell(++ro, co).Value = "Vertical = Center";
            ws.Cell(ro, co).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Cell(++ro, co).Value = "WrapText = true";
            ws.Cell(ro, co).Style.Alignment.WrapText = true;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}