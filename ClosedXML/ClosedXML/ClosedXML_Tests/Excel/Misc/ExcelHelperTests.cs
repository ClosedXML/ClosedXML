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
    public class ExcelHelperTests
    {

        [TestMethod]
        public void TestConvertColumnLetterToNumberAnd()
        {
            CheckColumnNumber(1);
            CheckColumnNumber(27);
            CheckColumnNumber(28);
            CheckColumnNumber(52);
            CheckColumnNumber(53);
            CheckColumnNumber(1000);
        }
        private static void CheckColumnNumber(int column)
        {
            Assert.AreEqual(column, ExcelHelper.GetColumnNumberFromLetter(ExcelHelper.GetColumnLetterFromNumber(column)));
        }
        
    }
}
