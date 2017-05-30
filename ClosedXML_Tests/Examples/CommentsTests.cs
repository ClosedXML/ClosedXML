using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class CommentsTests
    {
        [Test]
        public void AddingComments()
        {
            TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
        }
    }
}