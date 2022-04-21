using ClosedXML.Examples;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class PivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }
    }
}
