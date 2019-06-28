using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ClosedXML_Examples
{
    public class PivotTables : IXLExample
    {
        private class Pastry
        {
            public Pastry(string name, int? code, int numberOfOrders, double quality, string month, DateTime? bakeDate)
            {
                Name = name;
                Code = code;
                NumberOfOrders = numberOfOrders;
                Quality = quality;
                Month = month;
                BakeDate = bakeDate;
            }

            public string Name { get; set; }
            public int? Code { get; }
            public int NumberOfOrders { get; set; }
            public double Quality { get; set; }
            public string Month { get; set; }
            public DateTime? BakeDate { get; set; }
        }

        public void Create(String filePath)
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "Apr", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "May", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "Jun", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "Apr", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "May", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "Jun", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "Apr", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "May", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "Jun", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "Apr", new DateTime(2017, 04, 24)),
                new Pastry("Danish", 104, 190, 60, "May", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "Jun", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "Apr", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "May", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "Jun", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, null, null),
            };

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = ws.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                ws.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                #region Pivots

                for (int i = 1; i <= 3; i++)
                {
                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table.AsRange());

                    // The rows in our pivot table will be the names of the pastries
                    if (i == 2) pt.RowLabels.Add(XLConstants.PivotTableValuesSentinalLabel);
                    pt.RowLabels.Add("Name");

                    // The columns will be the months
                    pt.ColumnLabels.Add("Month");
                    if (i == 3) pt.ColumnLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                        .ShowAsPercentageFrom("Name").And("Bearclaw")
                        .NumberFormat.Format = "0%";

                    if (i > 1)
                    {
                        pt.Values.Add("Quality", "Sum of Quality")
                            .NumberFormat.SetFormat("#,##0.00");
                    }
                    if (i > 2)
                    {
                        pt.Values.Add("NumberOfOrders", "Sum of NumberOfOrders");
                    }

                    ptSheet.Columns().AdjustToContents();
                }

                #endregion Pivots

                #region Different kind of pivot

                ptSheet = wb.Worksheets.Add("pvtNoColumnLabels");
                pt = ptSheet.PivotTables.Add("pvtNoColumnLabels", ptSheet.Cell(1, 1), table.AsRange());

                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Month");

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                pt.SetRowHeaderCaption("Pastry name");

                #endregion Different kind of pivot

                #region Pivot table with collapsed fields

                ptSheet = wb.Worksheets.Add("pvtCollapsedFields");
                pt = ptSheet.PivotTables.Add("pvtCollapsedFields", ptSheet.Cell(1, 1), table.AsRange());

                pt.RowLabels.Add("Name").SetCollapsed();
                pt.RowLabels.Add("Month").SetCollapsed();

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                #endregion Pivot table with collapsed fields

                #region Pivot table with a field both as a value and as a row/column/filter label

                ptSheet = wb.Worksheets.Add("pvtFieldAsValueAndLabel");
                pt = ptSheet.PivotTables.Add("pvtFieldAsValueAndLabel", ptSheet.Cell(1, 1), table.AsRange());

                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Month");

                pt.Values.Add("Name").SetSummaryFormula(XLPivotSummary.Count);//.NumberFormat.Format = "#0.00";

                #endregion Pivot table with a field both as a value and as a row/column/filter label

                #region Pivot table with subtotals disabled

                ptSheet = wb.Worksheets.Add("pvtHideSubTotals");

                // Create the pivot table, using the data from the "PastrySalesData" table
                pt = ptSheet.PivotTables.Add("pvtHidesubTotals", ptSheet.Cell(1, 1), table.AsRange());

                // The rows in our pivot table will be the names of the pastries
                pt.RowLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

                // The columns will be the months
                pt.ColumnLabels.Add("Month");
                pt.ColumnLabels.Add("Name");

                // The values in our table will come from the "NumberOfOrders" field
                // The default calculation setting is a total of each row/column
                pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                    .ShowAsPercentageFrom("Name").And("Bearclaw")
                    .NumberFormat.Format = "0%";

                pt.Values.Add("Quality", "Sum of Quality")
                    .NumberFormat.SetFormat("#,##0.00");

                pt.Subtotals = XLPivotSubtotals.DoNotShow;

                pt.SetColumnHeaderCaption("Measures");

                ptSheet.Columns().AdjustToContents();

                #endregion Pivot table with subtotals disabled

                #region Pivot Table with filter

                ptSheet = wb.Worksheets.Add("pvtFilter");

                pt = table.CreatePivotTable(ptSheet.FirstCell(), "pvtFilter");

                pt.RowLabels.Add("Month");

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                pt.ReportFilters.Add("Name")
                    .AddSelectedValue("Scone")
                    .AddSelectedValue("Doughnut");

                pt.ReportFilters.Add("Quality")
                    .AddSelectedValue(5.19);

                pt.ReportFilters.Add("BakeDate")
                    .AddSelectedValue(new DateTime(2017, 05, 03));

                #endregion Pivot Table with filter

                #region Pivot table sorting

                ptSheet = wb.Worksheets.Add("pvtSort");
                pt = ptSheet.PivotTables.Add("pvtSort", ptSheet.Cell(1, 1), table.AsRange());

                pt.RowLabels.Add("Name").SetSort(XLPivotSortType.Ascending);
                pt.RowLabels.Add("Month").SetSort(XLPivotSortType.Descending);

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                pt.SetRowHeaderCaption("Pastry name");

                #endregion Pivot table sorting

                #region Pivot Table with integer rows

                ptSheet = wb.Worksheets.Add("pvtInteger");

                pt = ptSheet.PivotTables.Add("pvtInteger", ptSheet.Cell(1, 1), table);

                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Code");
                pt.RowLabels.Add("BakeDate");

                pt.ColumnLabels.Add("Month");

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                #endregion Pivot Table with integer rows

                wb.SaveAs(filePath);
            }
        }
    }
}
