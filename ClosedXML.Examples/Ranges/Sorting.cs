using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples.Misc
{
    public class Sorting : IXLExample
    {
        public void Create(String filePath)
        {
            using (var wb = new XLWorkbook())
            {

                #region Sort Table

                var wsTable = wb.Worksheets.Add("Table");
                AddTestTable(wsTable);

                wsTable.Row(1).InsertRowsAbove(1);
                Int32 lastCo = wsTable.LastColumnUsed().ColumnNumber();
                for (Int32 co = 1; co <= lastCo; co++)
                    wsTable.Cell(1, co).Value = "Column" + co.ToString();

                var table = wsTable.RangeUsed().AsTable();
                table.Sort("Column2 Desc, 1, 3 Asc");

                // Sort table another way
                wsTable = wb.Worksheets.Add("Table2");
                AddTestTable(wsTable);

                wsTable.Row(1).InsertRowsAbove(1);
                lastCo = wsTable.LastColumnUsed().ColumnNumber();
                for (Int32 co = 1; co <= lastCo; co++)
                    wsTable.Cell(1, co).Value = "Column" + co.ToString();

                table = wsTable.RangeUsed().AsTable();
                table.Sort("Column2", XLSortOrder.Descending, false, true);


                #endregion Sort Table

                #region Sort Rows

                var wsRows = wb.Worksheets.Add("Rows");
                AddTestTable(wsRows);
                wsRows.Row(1).Sort();
                wsRows.RangeUsed().Row(2).Sort();
                wsRows.Rows(3, wsRows.LastRowUsed().RowNumber()).Delete();

                #endregion Sort Rows

                #region Sort Columns

                var wsColumns = wb.Worksheets.Add("Columns");
                AddTestTable(wsColumns);
                wsColumns.LastColumnUsed().Delete();
                wsColumns.Column(1).Sort();
                wsColumns.RangeUsed().Column(2).Sort();

                #endregion Sort Columns

                #region Sort Mixed

                var wsMixed = wb.Worksheets.Add("Mixed");
                AddTestColumnMixed(wsMixed);
                wsMixed.Sort();

                #endregion Sort Mixed

                #region Sort Numbers

                var wsNumbers = wb.Worksheets.Add("Numbers");
                AddTestColumnNumbers(wsNumbers);
                wsNumbers.Sort();

                #endregion Sort Numbers

                #region Sort TimeSpans

                var wsTimeSpans = wb.Worksheets.Add("TimeSpans");
                AddTestColumnTimeSpans(wsTimeSpans);
                wsTimeSpans.Sort();

                #endregion Sort TimeSpans

                #region Sort Dates

                var wsDates = wb.Worksheets.Add("Dates");
                AddTestColumnDates(wsDates);
                wsDates.Sort();

                #endregion Sort Dates

                #region Do Not Ignore Blanks

                var wsIncludeBlanks = wb.Worksheets.Add("Include Blanks");
                AddTestTable(wsIncludeBlanks);
                var rangeIncludeBlanks = wsIncludeBlanks;
                rangeIncludeBlanks.SortColumns.Add(1, XLSortOrder.Ascending, false, true);
                rangeIncludeBlanks.SortColumns.Add(2, XLSortOrder.Descending, false, true);
                rangeIncludeBlanks.Sort();

                var wsIncludeBlanksColumn = wb.Worksheets.Add("Include Blanks Column");
                AddTestColumn(wsIncludeBlanksColumn);
                var rangeIncludeBlanksColumn = wsIncludeBlanksColumn;
                rangeIncludeBlanksColumn.SortColumns.Add(1, XLSortOrder.Ascending, false, true);
                rangeIncludeBlanksColumn.Sort();

                var wsIncludeBlanksColumnDesc = wb.Worksheets.Add("Include Blanks Column Desc");
                AddTestColumn(wsIncludeBlanksColumnDesc);
                var rangeIncludeBlanksColumnDesc = wsIncludeBlanksColumnDesc;
                rangeIncludeBlanksColumnDesc.SortColumns.Add(1, XLSortOrder.Descending, false, true);
                rangeIncludeBlanksColumnDesc.Sort();

                #endregion Do Not Ignore Blanks

                #region Case Sensitive

                var wsCaseSensitive = wb.Worksheets.Add("Case Sensitive");
                AddTestTable(wsCaseSensitive);
                var rangeCaseSensitive = wsCaseSensitive;
                rangeCaseSensitive.SortColumns.Add(1, XLSortOrder.Ascending, true, true);
                rangeCaseSensitive.SortColumns.Add(2, XLSortOrder.Descending, true, true);
                rangeCaseSensitive.Sort();

                var wsCaseSensitiveColumn = wb.Worksheets.Add("Case Sensitive Column");
                AddTestColumn(wsCaseSensitiveColumn);
                var rangeCaseSensitiveColumn = wsCaseSensitiveColumn;
                rangeCaseSensitiveColumn.SortColumns.Add(1, XLSortOrder.Ascending, true, true);
                rangeCaseSensitiveColumn.Sort();

                var wsCaseSensitiveColumnDesc = wb.Worksheets.Add("Case Sensitive Column Desc");
                AddTestColumn(wsCaseSensitiveColumnDesc);
                var rangeCaseSensitiveColumnDesc = wsCaseSensitiveColumnDesc;
                rangeCaseSensitiveColumnDesc.SortColumns.Add(1, XLSortOrder.Descending, true, true);
                rangeCaseSensitiveColumnDesc.Sort();

                #endregion Case Sensitive

                #region Simple Sorts

                var wsSimple = wb.Worksheets.Add("Simple");
                AddTestTable(wsSimple);
                wsSimple.Sort();

                var wsSimpleDesc = wb.Worksheets.Add("Simple Desc");
                AddTestTable(wsSimpleDesc);
                wsSimpleDesc.Sort("", XLSortOrder.Descending);

                var wsSimpleColumns = wb.Worksheets.Add("Simple Columns");
                AddTestTable(wsSimpleColumns);
                wsSimpleColumns.Sort("2, A DESC, 3");

                var wsSimpleColumn = wb.Worksheets.Add("Simple Column");
                AddTestColumn(wsSimpleColumn);
                wsSimpleColumn.Sort();

                var wsSimpleColumnDesc = wb.Worksheets.Add("Simple Column Desc");
                AddTestColumn(wsSimpleColumnDesc);
                wsSimpleColumnDesc.Sort(1, XLSortOrder.Descending);

                #endregion Simple Sorts

                wb.SaveAs(filePath);
            }
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
    }
}
