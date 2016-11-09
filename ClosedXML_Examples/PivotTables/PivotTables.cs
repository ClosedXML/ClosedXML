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
                new Pastry("Scone", 135, 0, "Apr"),
                new Pastry("Scone", 122, 5.19, "May"),
                new Pastry("Scone", 243, 44.2, "June")
            };

            using (var wb = new XLWorkbook())
            {

                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var source = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);

                // Create a range that includes our table, including the header row
                var range = source.DataRange;
                var header = sheet.Range(1, 1, 1, 3);
                var dataRange = sheet.Range(header.FirstCell(), range.LastCell());

                for (int i = 1; i <= 3; i++)
                {
                    // Add a new sheet for our pivot table
                    var ptSheet = wb.Worksheets.Add("PivotTable" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    var pt = ptSheet.PivotTables.AddNew("PivotTable", ptSheet.Cell(1, 1), dataRange);

                    // The rows in our pivot table will be the names of the pastries
                    pt.RowLabels.Add("Name");

                    // The columns will be the months
                    pt.ColumnLabels.Add("Month");

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("NumberOfOrders");
                }

                wb.SaveAs(filePath);
            }
        }
    }
}
