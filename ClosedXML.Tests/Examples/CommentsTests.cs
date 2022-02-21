using ClosedXML.Examples;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
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