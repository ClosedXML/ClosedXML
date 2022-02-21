using ClosedXML.Examples;
using ClosedXML.Examples.Columns;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class ColumnsTests
    {
        [Test]
        public void ColumnCells()
        {
            TestHelper.RunTestExample<ColumnCells>(@"Columns\ColumnCells.xlsx");
        }

        [Test]
        public void ColumnCollections()
        {
            TestHelper.RunTestExample<ColumnCollection>(@"Columns\ColumnCollection.xlsx");
        }

        [Test]
        public void ColumnSettings()
        {
            TestHelper.RunTestExample<ColumnSettings>(@"Columns\ColumnSettings.xlsx");
        }

        [Test]
        public void DeletingColumns()
        {
            TestHelper.RunTestExample<DeletingColumns>(@"Columns\DeletingColumns.xlsx");
        }

        //[Test] // Not working yet
        public void InsertColumns()
        {
            TestHelper.RunTestExample<InsertColumns>(@"Columns\InsertColumns.xlsx");
        }
    }
}
