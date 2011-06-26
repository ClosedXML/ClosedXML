using ClosedXML_Examples;
using ClosedXML_Examples.Rows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class RowsTests
    {
        [TestMethod]
        public void InsertRows()
        {
            TestHelper.RunTestExample<InsertRows>(@"Rows\InsertRows.xlsx");
        }
        [TestMethod]
        public void RowCells()
        {
            TestHelper.RunTestExample<RowCells>(@"Rows\RowCells.xlsx");
        }
        [TestMethod]
        public void RowCollection()
        {
            TestHelper.RunTestExample<RowCollection>(@"Rows\RowCollection.xlsx");
        }
        [TestMethod]
        public void RowSettings()
        {
            TestHelper.RunTestExample<RowSettings>(@"Rows\RowSettings.xlsx");
        }

    }
}