using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples.Misc
{
    public class DataTypes : IXLExample
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
            var ws = workbook.Worksheets.Add("Data Types");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "Plain Text:";
            ws.Cell(ro, co + 1).Value = "Hello World.";

            ws.Cell(++ro, co).Value = "Plain Date:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);

            ws.Cell(++ro, co).Value = "Plain DateTime:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);

            ws.Cell(++ro, co).Value = "Plain Boolean:";
            ws.Cell(ro, co + 1).Value = true;

            ws.Cell(++ro, co).Value = "Plain Number:";
            ws.Cell(ro, co + 1).Value = 123.45;

            ws.Cell(++ro, co).Value = "TimeSpan:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);

            ro++;

            ws.Cell(++ro, co).Value = "Decimal Number:";
            ws.Cell(ro, co + 1).Value = 123.45m;

            ws.Cell(++ro, co).Value = "Float Number:";
            ws.Cell(ro, co + 1).Value = 123.45f;

            ws.Cell(++ro, co).Value = "Double Number:";
            ws.Cell(ro, co + 1).Value = 123.45d;

            ro++;

            ws.Cell(++ro, co).Value = "Explicit Text:";
            ws.Cell(ro, co + 1).Value = "'Hello World.";

            ws.Cell(++ro, co).Value = "Date as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();

            ws.Cell(++ro, co).Value = "DateTime as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();

            ws.Cell(++ro, co).Value = "Boolean as Text:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();

            ws.Cell(++ro, co).Value = "Number as Text:";
            ws.Cell(ro, co + 1).Value = "'123.45";

            ws.Cell(++ro, co).Value = "Number with @ format:";
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "@";
            ws.Cell(ro, co + 1).Value = 123.45;

            ws.Cell(++ro, co).Value = "Format number with @:";
            ws.Cell(ro, co + 1).Value = 123.45;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "@";

            ws.Cell(++ro, co).Value = "TimeSpan as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();

            ro++;

            ws.Cell(++ro, co).Value = "Changing Data Types:";

            ro++;

            ws.Cell(++ro, co).Value = "Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "DateTime to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Boolean to Text:";
            ws.Cell(ro, co + 1).Value = true;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Number to Text:";
            ws.Cell(ro, co + 1).Value = 123.45;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "TimeSpan to Text:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Text to Date:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to DateTime:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to Boolean:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.Boolean;

            ws.Cell(++ro, co).Value = "Text to Number:";
            ws.Cell(ro, co + 1).Value = "'123.45";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Number;

            ws.Cell(++ro, co).Value = "@ format to Number:";
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "@";
            ws.Cell(ro, co + 1).Value = 123.45;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Number;

            ws.Cell(++ro, co).Value = "Text to TimeSpan:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.TimeSpan;

            ro++;

            ws.Cell(++ro, co).Value = "Formatted Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Formatted Number to Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ro++;

            ws.Cell(++ro, co).Value = "Blank Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;
            ws.Cell(ro, co + 1).Value = "";

            ro++;

            // Using inline strings (few users will ever need to use this feature)
            //
            // By default all strings are stored as shared so one block of text
            // can be reference by multiple cells.
            // You can override this by setting the .ShareString property to false
            ws.Cell(++ro, co).Value = "Inline String:";
            var cell = ws.Cell(ro, co + 1);
            cell.Value = "Not Shared";
            cell.ShareString = false;

            // To view all shared strings (all texts in the workbook actually), use the following:
            // workbook.GetSharedStrings()

            ws.Cell(++ro, co)
                .SetDataType(XLCellValues.Text)
                .SetDataType(XLCellValues.Boolean)
                .SetDataType(XLCellValues.DateTime)
                .SetDataType(XLCellValues.Number)
                .SetDataType(XLCellValues.TimeSpan)
                .SetDataType(XLCellValues.Text)
                .SetDataType(XLCellValues.TimeSpan)
                .SetDataType(XLCellValues.Number)
                .SetDataType(XLCellValues.DateTime)
                .SetDataType(XLCellValues.Boolean)
                .SetDataType(XLCellValues.Text);

            ws.Columns(2, 3).AdjustToContents();

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
