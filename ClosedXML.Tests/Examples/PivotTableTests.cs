using ClosedXML.Examples;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
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
