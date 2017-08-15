using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class PivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
        }
    }
}
