using ClosedXML_Examples;
using ClosedXML_Examples.Columns;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class ColumnsTests
    {
        [TestMethod]
        public void ColumnCells()
        {
            TestHelper.RunTestExample<ColumnCells>(@"Columns\ColumnCells.xlsx");
        }
        [TestMethod]
        public void ColumnCollections()
        {
            TestHelper.RunTestExample<ColumnCollection>(@"Columns\ColumnCollection.xlsx");
        }
        [TestMethod]
        public void ColumnSettings()
        {
            TestHelper.RunTestExample<ColumnSettings>(@"Columns\ColumnSettings.xlsx");
        }
        [TestMethod]
        public void DeletingColumns()
        {
            TestHelper.RunTestExample<DeletingColumns>(@"Columns\DeletingColumns.xlsx");
        }
        [TestMethod]
        public void InsertColumns()
        {
            TestHelper.RunTestExample<InsertColumns>(@"Columns\InsertColumns.xlsx");
        }

    }
}
