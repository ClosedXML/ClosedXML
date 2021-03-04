using ClosedXML.Examples;
using ClosedXML.Examples.Rows;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class RowsTests
    {
        [Test]
        public void RowCells()
        {
            TestHelper.RunTestExample<RowCells>(@"Rows\RowCells.xlsx");
        }

        [Test]
        public void RowCollection()
        {
            TestHelper.RunTestExample<RowCollection>(@"Rows\RowCollection.xlsx");
        }

        [Test]
        public void RowSettings()
        {
            TestHelper.RunTestExample<RowSettings>(@"Rows\RowSettings.xlsx");
        }

        //[Test] // Not working yet
        public void InsertRows()
        {
            TestHelper.RunTestExample<InsertRows>(@"Rows\InsertRows.xlsx");
        }
    }
}
