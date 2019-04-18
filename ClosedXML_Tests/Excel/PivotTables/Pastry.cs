using System;
using System.Linq;
using ClosedXML.Attributes;

namespace ClosedXML_Tests
{
    public class Pastry
    {
        public Pastry(string name, int? code, int numberOfOrders, double quality, double price, string month, DateTime? bakeDate)
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

        //public double Price { get; set; }
        // public double Summ => Quality * Price;
        public string Month { get; set; }
        public DateTime? BakeDate { get; set; }

        // Based on .\ClosedXML\ClosedXML_Examples\PivotTables\PivotTables.cs
        // But with empty column for Month
        [XLColumn(Ignore = true)] public static Pastry[] DefaultSet =
        {
            new Pastry("Croissant", 101, 150, 60.2, 1.8, "", new DateTime(2016, 04, 21)),
            new Pastry("Croissant", 101, 250, 50.42, 1.8, "", new DateTime(2016, 05, 03)),
            new Pastry("Croissant", 101, 134, 22.12, 1.82, "", new DateTime(2016, 06, 24)),
            new Pastry("Doughnut", 102, 250, 89.99, 1.09, "", new DateTime(2017, 04, 23)),
            new Pastry("Doughnut", 102, 225, 70, 1.09, "", new DateTime(2016, 05, 24)),
            new Pastry("Doughnut", 102, 210, 75.33, 1.09, "", new DateTime(2016, 06, 02)),
            new Pastry("Bearclaw", 103, 134, 10.24, 2.69, "", new DateTime(2016, 04, 27)),
            new Pastry("Bearclaw", 103, 184, 33.33, 2.69, "", new DateTime(2016, 05, 20)),
            new Pastry("Bearclaw", 103, 124, 25, 2.69, "", new DateTime(2017, 06, 05)),
            new Pastry("Danish", 104, 394, -20.24, 2.3, "", null),
            new Pastry("Danish", 104, 190, 60, 2.19, "", new DateTime(2017, 05, 08)),
            new Pastry("Danish", 104, 221, 24.76, 2.19, "", new DateTime(2016, 06, 21)),

            // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
            new Pastry("Scone", 105, 135, 0, 2.45, "", new DateTime(2017, 04, 22)),
            new Pastry("SconE", 105, 122, 5.19, 2.45, "", new DateTime(2017, 05, 03)),
            new Pastry("SCONE", 105, 243, 44.2, 2.45, "", new DateTime(2017, 06, 14)),

            // For ContainsBlank and integer rows/columns test
            new Pastry("Scone", null, 255, 18.4, 2.45, "", null),
        };

        [XLColumn(Ignore = true)]
        public static Pastry[] WithMonthSet
        {
            get
            {
                return DefaultSet.Select(x =>
                {
                    var clone = (Pastry)x.MemberwiseClone();
                    if (clone.BakeDate != null) clone.Month = clone.BakeDate.Value.ToString("MMM");
                    return clone;
                }).ToArray();
            }
        }
    }
}
