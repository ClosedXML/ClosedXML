using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples
{
    public class DeletingColumns
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
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Deleting Columns");

            var rngTitles = ws.Range("B2:D2");
            ws.Row(1).InsertRowsBelow(2);
            Console.Write(rngTitles.ToString()); // Prints "B4:D4
            Console.ReadKey();

            var rng1 = ws.Range("B2:D2"); 
            var rng2 = ws.Range("F2:G2");
            var rng3 = ws.Range("A1:A3");
            var col1 = ws.Column(1);

            // rng1 will have 2 columns starting at A2
            ws.Columns("A,C,E:H").Delete();

            rng1.Style.Fill.BackgroundColor = Color.Orange;
            rng2.Style.Fill.BackgroundColor = Color.Blue;
            //rng3.Style.Fill.BackgroundColor = Color.Red;
            //col1.Style.Fill.BackgroundColor = Color.Red;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
