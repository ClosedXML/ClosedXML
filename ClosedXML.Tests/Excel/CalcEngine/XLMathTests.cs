using ClosedXML.Excel.CalcEngine.Functions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class XLMathTests
    {
        [Test]
        public void IsEven()
        {
            Assert.IsTrue(XLMath.IsEven(2));
            Assert.IsFalse(XLMath.IsEven(3));
        }

        [Test]
        public void IsOdd()
        {
            Assert.IsTrue(XLMath.IsOdd(3));
            Assert.IsFalse(XLMath.IsOdd(2));
        }
    }
}
