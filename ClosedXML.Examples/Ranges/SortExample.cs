using System;
using System.Linq;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class SortExample : IXLExample
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

            #region Sort a table
            var wsTable = wb.Worksheets.Add("Table");
            AddTestTable(wsTable);
            var header = wsTable.Row(1).InsertRowsAbove(1).First();
            for(Int32 co = 1; co <= wsTable.LastColumnUsed().ColumnNumber(); co++)
            {
                header.Cell(co).Value = "Column" + co.ToString();
            }
            var rangeTable = wsTable.RangeUsed();
            var table = rangeTable.CopyTo(wsTable.Column(wsTable.LastColumnUsed().ColumnNumber() + 3)).CreateTable();

            table.Sort("Column2, Column3 Desc, Column1 ASC");
            
            wsTable.Row(1).InsertRowsAbove(2);
            wsTable.Cell(1, 1)
                .SetValue(".Sort(\"Column2, Column3 Desc, Column1 ASC\") = Sort table Top to Bottom, Col 2 Asc, Col 3 Desc, Col 1 Asc, Ignore Blanks, Ignore Case")
                .Style.Font.SetBold();
            #endregion

            #region Sort a simple range left to right
            var wsLeftToRight = wb.Worksheets.Add("Sort Left to Right");
            AddTestTable(wsLeftToRight);
            wsLeftToRight.RangeUsed().Transpose(XLTransposeOptions.MoveCells);
            var rangeLeftToRight = wsLeftToRight.RangeUsed();
            var copyLeftToRight = rangeLeftToRight.CopyTo(wsLeftToRight.Row(wsLeftToRight.LastRowUsed().RowNumber() + 3));

            copyLeftToRight.SortLeftToRight();

            wsLeftToRight.Row(1).InsertRowsAbove(2);
            wsLeftToRight.Cell(1, 1)
                .SetValue(".SortLeftToRight() = Sort Range Left to Right, Ascendingly, Ignore Blanks, Ignore Case")
                .Style.Font.SetBold();
            #endregion

            #region Sort a range
            var wsComplex2 = wb.Worksheets.Add("Complex 2");
            AddTestTable(wsComplex2);
            var rangeComplex2 = wsComplex2.RangeUsed();
            var copyComplex2 = rangeComplex2.CopyTo(wsComplex2.Column(wsComplex2.LastColumnUsed().ColumnNumber() + 3));

            copyComplex2.SortColumns.Add(1, XLSortOrder.Ascending, false, true);
            copyComplex2.SortColumns.Add(3, XLSortOrder.Descending);
            copyComplex2.Sort();

            wsComplex2.Row(1).InsertRowsAbove(4);
            wsComplex2.Cell(1, 1)
                .SetValue(".SortColumns.Add(1, XLSortOrder.Ascending, false, true) = Sort Col 1 Asc, Match Blanks, Match Case").Style.Font.SetBold();
            wsComplex2.Cell(2, 1)
                .SetValue(".SortColumns.Add(3, XLSortOrder.Descending) = Sort Col 3 Desc, Ignore Blanks, Ignore Case").Style.Font.SetBold();
            wsComplex2.Cell(3, 1)
                .SetValue(".Sort() = Sort range using the parameters defined in SortColumns").Style.Font.SetBold();
            #endregion

            #region Sort a range 
            var wsComplex1 = wb.Worksheets.Add("Complex 1");
            AddTestTable(wsComplex1);
            var rangeComplex1 = wsComplex1.RangeUsed();
            var copyComplex1 = rangeComplex1.CopyTo(wsComplex1.Column(wsComplex1.LastColumnUsed().ColumnNumber() + 3));

            copyComplex1.Sort("2, 1 DESC", XLSortOrder.Ascending, true);

            wsComplex1.Row(1).InsertRowsAbove(2);
            wsComplex1.Cell(1, 1)
                .SetValue(".Sort(\"2, 1 DESC\", XLSortOrder.Ascending, true) = Sort Range Top to Bottom, Col 2 Asc, Col 1 Desc, Ignore Blanks, Match Case").Style.Font.SetBold();
            #endregion

            #region Sort a simple column
            var wsSimpleColumn = wb.Worksheets.Add("Simple Column");
            AddTestColumn(wsSimpleColumn);
            var rangeSimpleColumn = wsSimpleColumn.RangeUsed();
            var copySimpleColumn = rangeSimpleColumn.CopyTo(wsSimpleColumn.Column(wsSimpleColumn.LastColumnUsed().ColumnNumber() + 3));

            copySimpleColumn.FirstColumn().Sort(XLSortOrder.Descending, true);

            wsSimpleColumn.Row(1).InsertRowsAbove(2);
            wsSimpleColumn.Cell(1, 1)
                .SetValue(".Sort(XLSortOrder.Descending, true) = Sort Range Top to Bottom, Descendingly, Ignore Blanks, Match Case").Style.Font.SetBold();
            #endregion

            #region Sort a simple range
            var wsSimple = wb.Worksheets.Add("Simple");
            AddTestTable(wsSimple);
            var rangeSimple = wsSimple.RangeUsed();
            var copySimple = rangeSimple.CopyTo(wsSimple.Column(wsSimple.LastColumnUsed().ColumnNumber() + 3));
            
            copySimple.Sort();

            wsSimple.Row(1).InsertRowsAbove(2);
            wsSimple.Cell(1, 1).SetValue(".Sort() = Sort Range Top to Bottom, Ascendingly, Ignore Blanks, Ignore Case").Style.Font.SetBold();
            #endregion

            wb.SaveAs(filePath);
        }

        private void AddTestColumnMixed(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue(new DateTime(2011, 1, 30)).Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue(1.15).Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue(new TimeSpan(1, 1, 12, 30)).Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue(9).Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue(new TimeSpan(9, 4, 30)).Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue(new DateTime(2011, 4, 15)).Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        private void AddTestColumnNumbers(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue(1.30).Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue(1.15).Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue(1230).Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue(9).Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue(4.30).Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue(4.15).Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        private void AddTestColumnTimeSpans(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue(new TimeSpan(0, 12, 35, 21)).Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue(new TimeSpan(45, 1, 15)).Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue(new TimeSpan(1, 1, 12, 30)).Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue(new TimeSpan(0, 12, 15)).Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue(new TimeSpan(1, 4, 30)).Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue(new TimeSpan(1, 4, 15)).Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        private void AddTestColumnDates(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue(new DateTime(2011, 1, 30)).Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue(new DateTime(2011, 1, 15)).Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue(new DateTime(2011, 12, 30)).Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue(new DateTime(2011, 12, 15)).Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue(new DateTime(2011, 4, 30)).Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue(new DateTime(2011, 4, 15)).Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        private void AddTestColumn(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue("B").Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue("b").Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue("B").Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue("c").Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        private void AddTestTable(IXLWorksheet ws)
        {
            ws.Cell("A1").SetValue("B").Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("A2").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("A3").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("A4").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("A5").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("A6").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("A7").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("A8").SetValue("B").Style.Fill.SetBackgroundColor(XLColor.DeepPink);

            ws.Cell("B1").SetValue("").Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("B2").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("B3").SetValue("B").Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("B4").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("B5").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("B6").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("B7").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("B8").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.DeepPink);

            ws.Cell("C1").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.LightGreen);
            ws.Cell("C2").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DarkTurquoise);
            ws.Cell("C3").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.BurlyWood);
            ws.Cell("C4").SetValue("a").Style.Fill.SetBackgroundColor(XLColor.DarkGray);
            ws.Cell("C5").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.DarkSalmon);
            ws.Cell("C6").SetValue("b").Style.Fill.SetBackgroundColor(XLColor.DodgerBlue);
            ws.Cell("C7").SetValue("A").Style.Fill.SetBackgroundColor(XLColor.IndianRed);
            ws.Cell("C8").SetValue("").Style.Fill.SetBackgroundColor(XLColor.DeepPink);
        }
        // Private

        // Override


        #endregion
    }
}
