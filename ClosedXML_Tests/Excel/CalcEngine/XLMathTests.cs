using ClosedXML.Excel.CalcEngine.Functions;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
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