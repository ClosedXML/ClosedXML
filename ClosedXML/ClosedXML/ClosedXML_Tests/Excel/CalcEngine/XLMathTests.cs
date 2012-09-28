using System;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class Extensions
    {

        [TestMethod]
        public void IsEven()
        {
            Assert.IsTrue(XLMath.IsEven(2));
            Assert.IsFalse(XLMath.IsEven(3));
        }

        [TestMethod]
        public void IsOdd()
        {
            Assert.IsTrue(XLMath.IsOdd(3));
            Assert.IsFalse(XLMath.IsOdd(2));
        }
    }
}
