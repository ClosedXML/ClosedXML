using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples.Misc
{
    public class FreezePanes : IXLExample
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
            var wsFreeze = wb.Worksheets.Add("Freeze View");
            
            // Freeze rows and columns in one shot
            wsFreeze.SheetView.Freeze(3, 3);

            // You can also be more specific on what you want to freeze
            // For example:
            // wsFreeze.SheetView.FreezeRows(3);
            // wsFreeze.SheetView.FreezeColumns(3);


            //////////////////////////////
            //var wsSplit = wb.Worksheets.Add("Split View");
            //wsSplit.SheetView.SplitRow = 3;
            //wsSplit.SheetView.SplitColumn = 3;

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
