using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.DataValidations
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class FunctionsTests
    {
        [TestMethod]
        public void Even()
        {
            var actual1 = XLWorkbook.EvaluateExpr("Even(1.5)");
            Assert.AreEqual(2.0, actual1);
            var actual2 = XLWorkbook.EvaluateExpr("Even(2.01)");
            Assert.AreEqual(4.0, actual2);
        }

        [TestMethod]
        public void Combin()
        {
            var actual1 = XLWorkbook.EvaluateExpr("Combin(200, 2)");
            Assert.AreEqual(19900.0, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("Combin(20.1, 2.9)");
            Assert.AreEqual(190.0, actual2);
        }

        [TestMethod]
        public void Degrees()
        {
            var actual1 = XLWorkbook.EvaluateExpr("Degrees(180)");
            Assert.IsTrue(Math.PI - (double)actual1 < XLHelper.Epsilon);
        }

        [TestMethod]
        public void Fact()
        {
            var actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
            Assert.AreEqual(120.0, actual);
        }

        [TestMethod]
        public void FactDouble()
        {
            var actual1 = XLWorkbook.EvaluateExpr("FactDouble(6)");
            Assert.AreEqual(48.0, actual1);
            var actual2 = XLWorkbook.EvaluateExpr("FactDouble(7)");
            Assert.AreEqual(105.0, actual2);
        }

        [TestMethod]
        public void Gcd()
        {
            var actual = XLWorkbook.EvaluateExpr("Gcd(24, 36)");
            Assert.AreEqual(12, actual);

            var actual1 = XLWorkbook.EvaluateExpr("Gcd(5, 0)");
            Assert.AreEqual(5, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("Gcd(0, 5)");
            Assert.AreEqual(5, actual2);

            var actual3 = XLWorkbook.EvaluateExpr("Gcd(240, 360, 30)");
            Assert.AreEqual(30, actual3);
        }

        [TestMethod]
        public void Lcm()
        {
            var actual = XLWorkbook.EvaluateExpr("Lcm(24, 36)");
            Assert.AreEqual(72, actual);

            var actual1 = XLWorkbook.EvaluateExpr("Lcm(5, 0)");
            Assert.AreEqual(0, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("Lcm(0, 5)");
            Assert.AreEqual(0, actual2);

            var actual3 = XLWorkbook.EvaluateExpr("Lcm(240, 360, 30)");
            Assert.AreEqual(720, actual3);
        }

        [TestMethod]
        public void Mod()
        {
            var actual = XLWorkbook.EvaluateExpr("Mod(3, 2)");
            Assert.AreEqual(1, actual);

            var actual1 = XLWorkbook.EvaluateExpr("Mod(-3, 2)");
            Assert.AreEqual(1, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("Mod(3, -2)");
            Assert.AreEqual(-1, actual2);

            var actual3 = XLWorkbook.EvaluateExpr("Mod(-3, -2)");
            Assert.AreEqual(-1, actual3);
        }

        [TestMethod]
        public void MRound()
        {
            //var actual = XLWorkbook.EvaluateExpr("MRound(10, 3)");
            //Assert.AreEqual(9.0, actual);

            //var actual1 = XLWorkbook.EvaluateExpr("MRound(-10, -10)");
            //Assert.AreEqual(-9.0, actual1);

            //var actual2 = XLWorkbook.EvaluateExpr("MRound(1.3, 0.2)");
            //Assert.AreEqual(1.4, actual2);
        }
    }
}
