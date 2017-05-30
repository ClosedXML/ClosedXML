using ClosedXML_Examples.Delete;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class DeleteTests
    {
        [Test]
        public void DeleteFewWorksheets()
        {
            TestHelper.RunTestExample<DeleteFewWorksheets>(@"Delete\DeleteFewWorksheets.xlsx");
        }

        [Test]
        public void RemoveRows()
        {
            TestHelper.RunTestExample<DeleteRows>(@"Delete\RemoveRows.xlsx");
        }
    }
}