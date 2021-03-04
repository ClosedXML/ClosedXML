using ClosedXML.Examples.Delete;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
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