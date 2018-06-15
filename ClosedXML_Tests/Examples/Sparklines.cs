using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class SparklinesTests
    {
        [Test]
        public void Sparklines()
        {
            TestHelper.RunTestExample<Sparklines>(@"Sparklines\Sparklines.xlsx");
        }
    }
}
