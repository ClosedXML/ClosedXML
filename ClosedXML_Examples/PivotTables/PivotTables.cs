using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ClosedXML_Examples
{
    public class PivotTables : IXLExample
    {
        private class Pastry
        {
            public Pastry(string name, int numberOfOrders, double quality, string month)
            {
                Name = name;
                NumberOfOrders = numberOfOrders;
                Quality = quality;
                Month = month;
            }

            public string Name { get; set; }
            public int NumberOfOrders { get; set; }
            public double Quality { get; set; }
            public string Month { get; set; }
        }

        public void Create(String filePath)
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 150, 60.2, "Apr"),
                new Pastry("Croissant", 250, 50.42, "May"),
                new Pastry("Croissant", 134, 22.12, "June"),
                new Pastry("Doughnut", 250, 89.99, "Apr"),
                new Pastry("Doughnut", 225, 70, "May"),
                new Pastry("Doughnut", 210, 75.33, "June"),
                new Pastry("Bearclaw", 134, 10.24, "Apr"),
                new Pastry("Bearclaw", 184, 33.33, "May"),
                new Pastry("Bearclaw", 124, 25, "June"),
                new Pastry("Danish", 394, -20.24, "Apr"),
                new Pastry("Danish", 190, 60, "May"),
                new Pastry("Danish", 221, 24.76, "June"),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 135, 0, "Apr"),
                new Pastry("SconE", 122, 5.19, "May"),
                new Pastry("SCONE", 243, 44.2, "June")
            };

            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var source = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Columns().AdjustToContents();

                // Create a range that includes our table, including the header row
                var range = source.DataRange;
                var header = sheet.Range(1, 1, 1, 3);
                var dataRange = sheet.Range(header.FirstCell(), range.LastCell());

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                for (int i = 1; i <= 3; i++)
                {
                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.AddNew("pvt", ptSheet.Cell(1, 1), dataRange);

                    // The rows in our pivot table will be the names of the pastries
                    pt.RowLabels.Add("Name");
                    if (i == 2) pt.RowLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

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

                // Different kind of pivot
                ptSheet = wb.Worksheets.Add("pvtNoColumnLabels");
                pt = ptSheet.PivotTables.AddNew("pvtNoColumnLabels", ptSheet.Cell(1, 1), dataRange);

                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Month");

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);


                // Pivot table with collapsed fields
                ptSheet = wb.Worksheets.Add("pvtCollapsedFields");
                pt = ptSheet.PivotTables.AddNew("pvtCollapsedFields", ptSheet.Cell(1, 1), dataRange);

                pt.RowLabels.Add("Name").SetCollapsed();
                pt.RowLabels.Add("Month").SetCollapsed();

                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);


                // Pivot table with a field both as a value and as a row/column/filter label
                ptSheet = wb.Worksheets.Add("pvtFieldAsValueAndLabel");
                pt = ptSheet.PivotTables.AddNew("pvtFieldAsValueAndLabel", ptSheet.Cell(1, 1), dataRange);

                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Month");

                pt.Values.Add("Name").SetSummaryFormula(XLPivotSummary.Count);//.NumberFormat.Format = "#0.00";

                wb.SaveAs(filePath);
            }
        }
    }
}
