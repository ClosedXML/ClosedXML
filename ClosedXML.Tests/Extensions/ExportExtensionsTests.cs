using System.Collections.Generic;
using ClosedXML.Attributes;
using ClosedXML.Excel;
using ClosedXML.Extensions;
using NUnit.Framework;

namespace ClosedXML.Tests.Extensions
{
    public class ExportExtensionsTests
    {
        [Test]
        public void EasyExport_WithData_ReturnsXLWorkbook()
        {
            // Arrange
            var data = new List<IXLExportable>
            {
                new SampleXLExportable { FirstName = "John", LastName = "Doe", Age = 17, Email = "john@example.com", PhoneNumber = "1234567890" },
                new SampleXLExportable { FirstName = "Jane", LastName = "Smith", Age = 25, Email = "jane@example.com", PhoneNumber = "9876543210" },
                new SampleXLExportable { FirstName = "Michael", LastName = "Johnson", Age = 40, Email = "michael@example.com", PhoneNumber = "1112223333" },
                new SampleXLExportable { FirstName = "Emily", LastName = "Davis", Age = 35, Email = "emily@example.com", PhoneNumber = "4445556666" },
                new SampleXLExportable { FirstName = "David", LastName = "Brown", Age = 15, Email = "david@example.com", PhoneNumber = "7778889999" }
            };

            var options = new XLExportOptions
            {
                SheetName = "SampleSheet"
            };

            // Act
            var workbook = data.EasyExport<SampleXLExportable>(options, rowCallback: (field, obj, _) =>
            {
                switch (field.Property.Name)
                {
                    case nameof(SampleXLExportable.Age):
                    {
                        var cellResult = new XLExportDrawCellResult();

                        var ageCategory = obj.Age < 18 ? "Minor" : "Adult";
                        var textColor = obj.Age < 18 ? XLColor.Red : XLColor.Green;

                        cellResult.Value = ageCategory;
                        cellResult.Options = new XLExportCellOptions
                        {
                            TextColor = textColor,
                            FontSize = 12,
                            Bold = true
                        };

                        return cellResult;
                    }
                }

                return null;
            });

            workbook.SaveAs("sampleFile.xlsx");

            // Assert
            Assert.NotNull(workbook);
        }

        // Sample implementation of IXLExportable for testing purposes
        private class SampleXLExportable : IXLExportable
        {
            [XLColumn(Header = "First name", Order = 1)]
            public string FirstName { get; init; }

            [XLColumn(Header = "Last name", Order = 2)]
            public string LastName { get; init; }

            [XLColumn(Header = "Email address", Order = 4)]
            public string Email { get; init; }

            [XLColumn(Ignore = true)]
            public string PhoneNumber { get; init; }

            [XLColumn(Order = 3)]
            public int Age { get; init; }
        }
    }
}
