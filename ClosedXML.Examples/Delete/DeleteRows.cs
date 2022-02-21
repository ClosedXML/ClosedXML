using System;
using System.Linq;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Delete
{
    public class DeleteRows : IXLExample
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
            #region Create case
            {
                var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Delete red rows");

                // Put a value in a few cells
                foreach (var r in Enumerable.Range(1, 5))
                    foreach (var c in Enumerable.Range(1, 5))
                        ws.Cell(r, c).Value = string.Format("R{0}C{1}", r, c);


                var blueRow = ws.Rows(1, 2);
                var redRow = ws.Row(5);

                blueRow.Style.Fill.BackgroundColor = XLColor.Blue;

                redRow.Style.Fill.BackgroundColor = XLColor.Red;
                workbook.SaveAs(filePath);
            }
            #endregion

            #region Remove rows
            {
                var workbook = new XLWorkbook(filePath);
                var ws = workbook.Worksheets.Worksheet("Delete red rows");

                ws.Rows(1, 2).Delete();
                workbook.Save();
            }
            #endregion

        }

        // Private

        // Override


        #endregion
    }
}
