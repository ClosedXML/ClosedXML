using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class Extensions
    {

        [TestMethod]
        public void FixNewLines()
        {
            Assert.AreEqual("\n".FixNewLines(), Environment.NewLine);
            Assert.AreEqual("\r\n".FixNewLines(), Environment.NewLine);
            Assert.AreEqual("\rS\n".FixNewLines(), "\rS" + Environment.NewLine);
            Assert.AreEqual("\r\n\n".FixNewLines(), Environment.NewLine + Environment.NewLine);
        }

        
    }
}
