using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples.Misc
{
    public class CellValues : IXLExample
    {
        public void Create(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Cell Values");

                // Set the titles
                ws.Cell(2, 2).Value = "Initial Value";
                ws.Cell(2, 3).Value = "Casting";
                ws.Cell(2, 4).Value = "Using Get...()";
                ws.Cell(2, 5).Value = "Using GetValue<T>()";
                ws.Cell(2, 6).Value = "GetString()";
                ws.Cell(2, 7).Value = "GetFormattedString()";

                //////////////////////////////////////////////////////////////////
                // DateTime

                // Fill a cell with a date
                var cellDateTime = ws.Cell(3, 2);
                cellDateTime.Value = new DateTime(2010, 9, 2);
                cellDateTime.Style.DateFormat.Format = "yyyy-MMM-dd";

                // Extract the date in different ways
                DateTime dateTime1 = (DateTime)cellDateTime.Value;
                DateTime dateTime2 = cellDateTime.GetDateTime();
                DateTime dateTime3 = cellDateTime.GetValue<DateTime>();
                string dateTimeString = cellDateTime.GetString();
                string dateTimeFormattedString = cellDateTime.GetFormattedString();

                // Set the values back to cells
                // The apostrophe is to force ClosedXML to treat the date as a string
                ws.Cell(3, 3).Value = dateTime1;
                ws.Cell(3, 4).Value = dateTime2;
                ws.Cell(3, 5).Value = dateTime3;
                ws.Cell(3, 6).Value = "'" + dateTimeString;
                ws.Cell(3, 7).Value = "'" + dateTimeFormattedString;

                //////////////////////////////////////////////////////////////////
                // Boolean

                // Fill a cell with a boolean
                var cellBoolean = ws.Cell(4, 2);
                cellBoolean.Value = true;

                // Extract the boolean in different ways
                bool boolean1 = (bool)cellBoolean.Value;
                bool boolean2 = cellBoolean.GetBoolean();
                bool boolean3 = cellBoolean.GetValue<bool>();
                string booleanString = cellBoolean.GetString();
                string booleanFormattedString = cellBoolean.GetFormattedString();

                // Set the values back to cells
                // The apostrophe is to force ClosedXML to treat the boolean as a string
                ws.Cell(4, 3).Value = boolean1;
                ws.Cell(4, 4).Value = boolean2;
                ws.Cell(4, 5).Value = boolean3;
                ws.Cell(4, 6).Value = "'" + booleanString;
                ws.Cell(4, 7).Value = "'" + booleanFormattedString;

                //////////////////////////////////////////////////////////////////
                // Double

                // Fill a cell with a double
                var cellDouble = ws.Cell(5, 2);
                cellDouble.Value = 1234.567;
                cellDouble.Style.NumberFormat.Format = "#,##0.00";

                // Extract the double in different ways
                double double1 = (double)cellDouble.Value;
                double double2 = cellDouble.GetDouble();
                double double3 = cellDouble.GetValue<double>();
                string doubleString = cellDouble.GetString();
                string doubleFormattedString = cellDouble.GetFormattedString();

                // Set the values back to cells
                // The apostrophe is to force ClosedXML to treat the double as a string
                ws.Cell(5, 3).Value = double1;
                ws.Cell(5, 4).Value = double2;
                ws.Cell(5, 5).Value = double3;
                ws.Cell(5, 6).Value = "'" + doubleString;
                ws.Cell(5, 7).Value = "'" + doubleFormattedString;

                //////////////////////////////////////////////////////////////////
                // String

                // Fill a cell with a string
                var cellString = ws.Cell(6, 2);
                cellString.Value = "Test Case";

                // Extract the string in different ways
                string string1 = (string)cellString.Value;
                string string2 = cellString.GetString();
                string string3 = cellString.GetValue<string>();
                string stringString = cellString.GetString();
                string stringFormattedString = cellString.GetFormattedString();

                // Set the values back to cells
                ws.Cell(6, 3).Value = string1;
                ws.Cell(6, 4).Value = string2;
                ws.Cell(6, 5).Value = string3;
                ws.Cell(6, 6).Value = stringString;
                ws.Cell(6, 7).Value = stringFormattedString;

                //////////////////////////////////////////////////////////////////
                // TimeSpan

                // Fill a cell with a timeSpan
                var cellTimeSpan = ws.Cell(7, 2);
                cellTimeSpan.Value = new TimeSpan(1, 2, 31, 45);

                // Extract the timeSpan in different ways
                TimeSpan timeSpan1 = (TimeSpan)cellTimeSpan.Value;
                TimeSpan timeSpan2 = cellTimeSpan.GetTimeSpan();
                TimeSpan timeSpan3 = cellTimeSpan.GetValue<TimeSpan>();
                string timeSpanString = "'" + cellTimeSpan.GetString();
                string timeSpanFormattedString = "'" + cellTimeSpan.GetFormattedString();

                // Set the values back to cells
                ws.Cell(7, 3).Value = timeSpan1;
                ws.Cell(7, 4).Value = timeSpan2;
                ws.Cell(7, 5).Value = timeSpan3;
                ws.Cell(7, 6).Value = timeSpanString;
                ws.Cell(7, 7).Value = timeSpanFormattedString;

                //////////////////////////////////////////////////////////////////
                // Do some formatting
                ws.Columns("B:G").Width = 20;
                var rngTitle = ws.Range("B2:G2");
                rngTitle.Style.Font.Bold = true;
                rngTitle.Style.Fill.BackgroundColor = XLColor.Cyan;

                ws.Columns().AdjustToContents();

                ws = workbook.AddWorksheet("Test Whitespace");
                ws.FirstCell().Value = "'    ";

                workbook.SaveAs(filePath);
            }
        }
    }
}
