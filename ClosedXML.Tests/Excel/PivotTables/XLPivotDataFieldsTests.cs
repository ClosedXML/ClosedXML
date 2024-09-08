using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables
{
    /// <summary>
    /// Test methods of interface <see cref="IXLPivotValues"/> implemented through <see cref="XLPivotDataFields"/> class.
    /// </summary>
    internal class XLPivotDataFieldsTests
    {
        #region IXLPivotValues methods

        #region Add

        [Test]
        public void Add_source_name_must_be_from_pivot_cache_field_names()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("Name", "Price"),
                ("Cake", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);

            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => pt.Values.Add("Wrong field name"));

            Assert.NotNull(ex);
            StringAssert.StartsWith("Field 'Wrong field name' is not in the fields of a pivot cache. Should be one of 'Name','Price'.", ex.Message);
        }

        #endregion

        #endregion
    }
}
