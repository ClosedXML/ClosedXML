using ClosedXML_Examples.Delete;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class DeleteTests
    {
        [TestMethod]
        public void RemoveRows()
        {
            TestHelper.RunTestExample<DeleteRows>(@"Delete\RemoveRows.xlsx");
        }

    }
}