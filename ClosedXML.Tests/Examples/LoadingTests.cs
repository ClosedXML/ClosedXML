using ClosedXML.Examples;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class LoadingTests
    {
        [Test]
        public void ChangingBasicTable()
        {
            TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }
    }
}
