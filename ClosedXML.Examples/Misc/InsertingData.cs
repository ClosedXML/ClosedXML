using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class InsertingData : IXLExample
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Inserting Data");

                // From a list of strings
                var listOfStrings = new List<String>();
                listOfStrings.Add("House");
                listOfStrings.Add("001");
                ws.Cell(1, 1).Value = "From Strings";
                ws.Cell(1, 1).AsRange().AddToNamed("Titles");
                ws.Cell(2, 1).InsertData(listOfStrings);

                // From a list of arrays
                var listOfArr = new List<Int32[]>();
                listOfArr.Add(new Int32[] { 1, 2, 3 });
                listOfArr.Add(new Int32[] { 1 });
                listOfArr.Add(new Int32[] { 1, 2, 3, 4, 5, 6 });
                ws.Cell(1, 3).Value = "From Arrays";
                ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
                ws.Cell(2, 3).InsertData(listOfArr);

                // From a DataTable
                var dataTable = GetTable();
                ws.Cell(6, 1).Value = "From DataTable";
                ws.Range(6, 1, 6, 4).Merge().AddToNamed("Titles");
                ws.Cell(7, 1).InsertData(dataTable);

                // From a query
                var list = new List<Person>();
                list.Add(new Person() { Name = "John", Age = 30, House = "On Elm St." });
                list.Add(new Person() { Name = "Mary", Age = 15, House = "On Main St." });
                list.Add(new Person() { Name = "Luis", Age = 21, House = "On 23rd St." });
                list.Add(new Person() { Name = "Henry", Age = 45, House = "On 5th Ave." });

                var people = from p in list
                             where p.Age >= 21
                             select new { p.Name, p.House, p.Age };

                ws.Cell(6, 6).Value = "From Query";
                ws.Range(6, 6, 6, 8).Merge().AddToNamed("Titles");
                ws.Cell(7, 6).InsertData(people);

                ws.Cell(11, 6).Value = "From List";
                ws.Range(11, 6, 11, 9).Merge().AddToNamed("Titles");
                ws.Cell(12, 6).InsertData(list);

                ws.Cell("A13").Value = "Transposed";
                ws.Range(13, 1, 13, 3).Merge().AddToNamed("Titles");
                ws.Cell("A14").InsertData(people.AsEnumerable(), true);

                // Prepare the style for the titles
                var titlesStyle = wb.Style;
                titlesStyle.Font.Bold = true;
                titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

                // Format all titles in one shot
                wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

                ws.Columns().AdjustToContents();

                wb.SaveAs(filePath);
            }
        }

        private class Person
        {
            public String House { get; set; }
            public String Name { get; set; }
            public Int32 Age { get; set; }
            public static String ClassType { get { return nameof(Person); } }
        }

        // Private
        private DataTable GetTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            return table;
        }

        // Override

        #endregion Methods
    }
}
