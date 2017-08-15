using ClosedXML_Examples;
using ClosedXML_Examples.Rows;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
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
    }
}