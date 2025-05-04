using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables
{
    /// <summary>
    /// Test methods of interface <see cref="IXLPivotFields"/> implemented through <see cref="XLPivotTableAxis"/>.
    /// </summary>
    [TestFixture]
    internal class XLPivotTableAxisTests
    {
        #region IXLPivotFields methods

        #region Add

        [Test]
        public void Add_field_not_yet_in_table_adds_field_and_shared_items()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Count"),
                (1, 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            var internalPt = (XLPivotTable)pt;
            Assert.IsEmpty(internalPt.PivotFields[0].Items);

            var idField = pt.RowLabels.Add("ID", "Item ID").AddSubtotal(XLSubtotalFunction.Automatic);

            Assert.Multiple(() =>
            {
                Assert.That(idField.SourceName, Is.EqualTo("ID"));
                Assert.That(idField.CustomName, Is.EqualTo("Item ID"));
                Assert.That(pt.RowLabels.Single().CustomName, Is.EqualTo("Item ID"));
            });

            // Adds values and default aggregation func to items of the field
            var fieldItems = internalPt.PivotFields[0].Items;
            Assert.That(fieldItems, Has.Count.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(fieldItems[0].ItemType, Is.EqualTo(XLPivotItemType.Data));
                Assert.That(fieldItems[0].ItemIndex, Is.EqualTo(0));
                Assert.That(fieldItems[1].ItemType, Is.EqualTo(XLPivotItemType.Default));
            });
        }

        [Test]
        public void Same_field_cant_be_added_twice_to_same_axis()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Count"),
                (1, 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            pt.RowLabels.Add("ID", "Item ID");

            var ex = Assert.Throws<InvalidOperationException>(() => pt.RowLabels.Add("ID", "Item ID"))!;
            Assert.That(ex.Message, Is.EqualTo("Custom name 'Item ID' is already used."));
        }

        [Test]
        public void Add_field_must_exist_in_cache()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Count"),
                (1, 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            Assert.DoesNotThrow(() => pt.RowLabels.Add("ID", "Item ID"));

            var ex = Assert.Throws<InvalidOperationException>(() => pt.RowLabels.Add("nonexistent"))!;
            Assert.That(ex.Message, Is.EqualTo("Field 'nonexistent' not found in pivot cache."));
        }

        #endregion

        #region Clear

        [Test]
        public void Clear_removes_all_fields_from_axis()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            pt.RowLabels.Add("ID", "Item ID");
            pt.RowLabels.Add("Color", "Custom color");

            pt.RowLabels.Clear();

            Assert.IsEmpty(pt.RowLabels);

            // Clear should also remove custom names and axis, otherwise there are problems loading
            // file with such remains in Excel.
            var internalPt = (XLPivotTable)pt;
            Assert.Null(internalPt.PivotFields[0].Name);
            Assert.Null(internalPt.PivotFields[0].Axis);
            Assert.Null(internalPt.PivotFields[1].Name);
            Assert.Null(internalPt.PivotFields[1].Axis);
        }

        #endregion

        #region Contains

        [Test]
        public void Contains_checks_whether_field_is_present()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            var idField = pt.RowLabels.Add("ID", "Item ID");
            pt.ColumnLabels.Add("Color");

            Assert.Multiple(() =>
            {
                Assert.That(pt.RowLabels.Contains("id"), Is.True);
                Assert.That(pt.RowLabels.Contains(idField), Is.True);
            });
            Assert.False(pt.RowLabels.Contains("color"));
            Assert.False(pt.RowLabels.Contains("nonexistent"));
        }

        #endregion

        #region Get(string sourceName)

        [Test]
        public void Get_field_by_source_name()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            pt.RowLabels.Add("ID", "Item ID");
            pt.ColumnLabels.Add("Color");

            Assert.That(pt.RowLabels.Get("id").SourceName, Is.EqualTo("ID"));
            var ex = Assert.Throws<KeyNotFoundException>(() => pt.RowLabels.Get("color"))!;
            Assert.That(ex.Message, Is.EqualTo("Field with source name 'color' not found in AxisRow."));
        }

        #endregion

        #region Get(int)

        [Test]
        public void Get_field_by_index()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            pt.RowLabels.Add("ID", "Item ID");
            pt.ColumnLabels.Add("Color");

            Assert.That(pt.RowLabels.Get(0).SourceName, Is.EqualTo("ID"));
            Assert.Throws<IndexOutOfRangeException>(() => pt.RowLabels.Get(-2));
            Assert.Throws<IndexOutOfRangeException>(() => pt.RowLabels.Get(1));
        }

        #endregion

        #region IndexOf

        [Test]
        public void IndexOf_finds_field_in_axis_by_source_name()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            var idField = pt.RowLabels.Add("ID", "Item ID");
            pt.ColumnLabels.Add("Color");

            Assert.Multiple(() =>
            {
                Assert.That(pt.RowLabels.IndexOf("ID"), Is.EqualTo(0));
                Assert.That(pt.RowLabels.IndexOf(idField), Is.EqualTo(0));
                Assert.That(pt.RowLabels.IndexOf("item id"), Is.EqualTo(-1));
                Assert.That(pt.RowLabels.IndexOf("Color"), Is.EqualTo(-1));
            });
        }

        #endregion

        #region Remove

        [Test]
        public void Remove_removes_field()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet();
            var range = data.Cell("A1").InsertData(new object[]
            {
                ("ID", "Color", "Count"),
                (1, "Blue", 10),
            });
            var ptSheet = wb.AddWorksheet();
            var pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
            pt.RowLabels.Add("ID");

            pt.RowLabels.Remove("id");
            pt.RowLabels.Remove("ID"); // Doesnt throw on already removed.

            Assert.IsEmpty(pt.RowLabels);
        }

        #endregion

        #endregion
    }
}
