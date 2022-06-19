using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class Outline : IXLExample
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

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Outline");

            ws.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Right;
            ws.Columns(2, 6).Group(); // Create an outline (level 1) for columns 2-6
            ws.Columns(2, 4).Group(); // Create an outline (level 2) for columns 2-4
            ws.Column(2).Ungroup(true); // Remove column 2 from all outlines

            ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Bottom;
            ws.Rows(1, 5).Group(); // Create an outline (level 1) for rows 1-5
            ws.Rows(1, 4).Group(); // Create an outline (level 2) for rows 1-4
            ws.Rows(1, 4).Collapse(); // Collapse rows 1-4
            ws.Rows(1, 2).Group(); // Create an outline (level 3) for rows 1-2
            ws.Rows(1, 2).Ungroup(); // Ungroup rows 1-2 from their last outline

            // You can also Collapse/Expand specific outline levels
            // 
            // ws.CollapseRows(Int32 outlineLevel)
            // ws.CollapseColumns(Int32 outlineLevel)
            //
            // ws.ExpandRows(Int32 outlineLevel)
            // ws.ExpandColumns(Int32 outlineLevel)

            // And you can also Collapse/Expand ALL outline levels in one shot
            // 
            // ws.CollapseRows()
            // ws.CollapseColumns()
            //
            // ws.ExpandRows()
            // ws.ExpandColumns()

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}