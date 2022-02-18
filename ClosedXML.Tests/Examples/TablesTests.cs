using ClosedXML.Examples.Tables;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class TablesTests
    {
        [Test]
        public void InsertingTables()
        {
            var allowedDiff = "/xl/worksheets/sheet1.xml :NonEqual\n";

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                allowedDiff = null;
            }

            TestHelper.RunTestExample<InsertingTables>(@"Tables\InsertingTables.xlsx", false, allowedDiff);
        }

        [Test]
        [Ignore("Don't know, why this new test of Francois Botha fails")]
        public void ResizingTables()
        {
            TestHelper.RunTestExample<ResizingTables>(@"Tables\ResizingTables.xlsx");
        }

        [Test]
        public void UsingTables()
        {
            TestHelper.RunTestExample<UsingTables>(@"Tables\UsingTables.xlsx");
        }
    }
}
