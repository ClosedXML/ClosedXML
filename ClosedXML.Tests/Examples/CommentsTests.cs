using ClosedXML.Examples;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class CommentsTests
    {
        [Test]
        public void AddingComments()
        {
            TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }
    }
}
