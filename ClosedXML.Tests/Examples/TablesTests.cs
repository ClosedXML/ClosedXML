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
            TestHelper.RunTestExample<InsertingTables>(@"Tables\InsertingTables.xlsx", false, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void ResizingTables()
        {
            TestHelper.RunTestExample<ResizingTables>(@"Tables\ResizingTables.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void UsingTables()
        {
            TestHelper.RunTestExample<UsingTables>(@"Tables\UsingTables.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }
    }
}
