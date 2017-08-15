using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class LoadingTests
    {
        [Test]
        public void ChangingBasicTable()
        {
            TestHelper.RunTestExample<ChangingBasicTable>(@"Loading\ChangingBasicTable.xlsx");
        }
    }
}