using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
            var actual = XLWorkbook.EvaluateExpr("MRound(10, 3)");
            Assert.AreEqual(9m, actual);

            var actual3 = XLWorkbook.EvaluateExpr("MRound(10.5, 3)");
            Assert.AreEqual(12m, actual3);

            var actual4 = XLWorkbook.EvaluateExpr("MRound(10.4, 3)");
            Assert.AreEqual(9m, actual4);

            var actual1 = XLWorkbook.EvaluateExpr("MRound(-10, -3)");
            Assert.AreEqual(-9m, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("MRound(1.3, 0.2)");
            Assert.AreEqual(1.4m, actual2);
        }

        [TestMethod]
        public void Multinomial()
        {
            var actual = XLWorkbook.EvaluateExpr("Multinomial(2,3,4)");
            Assert.AreEqual(1260.0, actual);
        }

        [TestMethod]
        public void Odd()
        {
            var actual = XLWorkbook.EvaluateExpr("Odd(1.5)");
            Assert.AreEqual(3, actual);

            var actual1 = XLWorkbook.EvaluateExpr("Odd(3)");
            Assert.AreEqual(3, actual1);

            var actual2 = XLWorkbook.EvaluateExpr("Odd(2)");
            Assert.AreEqual(3, actual2);

            var actual3 = XLWorkbook.EvaluateExpr("Odd(-1)");
            Assert.AreEqual(-1, actual3);

            var actual4 = XLWorkbook.EvaluateExpr("Odd(-2)");
            Assert.AreEqual(-3, actual4);

            actual = XLWorkbook.EvaluateExpr("Odd(0)");
            Assert.AreEqual(1, actual);
        }

        [TestMethod]
        public void Even()
        {
            var actual = XLWorkbook.EvaluateExpr("Even(3)");
            Assert.AreEqual(4, actual);

            actual = XLWorkbook.EvaluateExpr("Even(2)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(-1)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(-2)");
            Assert.AreEqual(-2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(0)");
            Assert.AreEqual(0, actual);

            actual = XLWorkbook.EvaluateExpr("Even(1.5)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Even(2.01)");
            Assert.AreEqual(4, actual);
        }

        [TestMethod]
        public void Product()
        {
            var actual = XLWorkbook.EvaluateExpr("Product(2,3,4)");
            Assert.AreEqual(24.0, actual);
        }

        [TestMethod]
        public void Quotient()
        {
            var actual = XLWorkbook.EvaluateExpr("Quotient(5,2)");
            Assert.AreEqual(2, actual);

            actual = XLWorkbook.EvaluateExpr("Quotient(4.5,3.1)");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr("Quotient(-10,3)");
            Assert.AreEqual(-3, actual);
        }

        [TestMethod]
        public void Radians()
        {
            var actual = XLWorkbook.EvaluateExpr("Radians(270)");
            Assert.IsTrue(Math.Abs(4.71238898038469 - (double)actual) < XLHelper.Epsilon);
        }

        [TestMethod]
        public void Roman()
        {
            var actual = XLWorkbook.EvaluateExpr("Roman(3046, 1)");
            Assert.AreEqual("MMMXLVI", actual);

            actual = XLWorkbook.EvaluateExpr("Roman(270)");
            Assert.AreEqual("CCLXX", actual);

            actual = XLWorkbook.EvaluateExpr("Roman(3999, true)");
            Assert.AreEqual("MMMCMXCIX", actual);
        }

        [TestMethod]
        public void Round()
        {
            var actual = XLWorkbook.EvaluateExpr("Round(2.15, 1)");
            Assert.AreEqual(2.2, actual);

            actual = XLWorkbook.EvaluateExpr("Round(2.149, 1)");
            Assert.AreEqual(2.1, actual);

            actual = XLWorkbook.EvaluateExpr("Round(-1.475, 2)");
            Assert.AreEqual(-1.48, actual);

            actual = XLWorkbook.EvaluateExpr("Round(21.5, -1)");
            Assert.AreEqual(20.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(626.3, -3)");
            Assert.AreEqual(1000.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(1.98, -1)");
            Assert.AreEqual(0.0, actual);

            actual = XLWorkbook.EvaluateExpr("Round(-50.55, -2)");
            Assert.AreEqual(-100.0, actual);
        }

        [TestMethod]
        public void RoundDown()
        {
            var actual = XLWorkbook.EvaluateExpr("RoundDown(3.2, 0)");
            Assert.AreEqual(3.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(76.9, 0)");
            Assert.AreEqual(76.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(3.14159, 3)");
            Assert.AreEqual(3.141, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(-3.14159, 1)");
            Assert.AreEqual(-3.1, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(31415.92654, -2)");
            Assert.AreEqual(31400.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundDown(0, 3)");
            Assert.AreEqual(0.0, actual);
        }

        [TestMethod]
        public void RoundUp()
        {
            var actual = XLWorkbook.EvaluateExpr("RoundUp(3.2, 0)");
            Assert.AreEqual(4.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(76.9, 0)");
            Assert.AreEqual(77.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(3.14159, 3)");
            Assert.AreEqual(3.142, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(-3.14159, 1)");
            Assert.AreEqual(-3.2, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(31415.92654, -2)");
            Assert.AreEqual(31500.0, actual);

            actual = XLWorkbook.EvaluateExpr("RoundUp(0, 3)");
            Assert.AreEqual(0.0, actual);
        }

        [TestMethod]
        public void SeriesSum()
        {
            var actual = XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)");
            Assert.AreEqual(40.0, actual);

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A2").FormulaA1 = "PI()/4";
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
            ws.Cell("A5").FormulaA1 = "1/FACT(4)";
            ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

            actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
            Assert.IsTrue(Math.Abs(0.70710321482284566 - (double)actual) < XLHelper.Epsilon);
        }

        [TestMethod]
        public void SqrtPi()
        {
            var actual = XLWorkbook.EvaluateExpr("SqrtPi(1)");
            Assert.IsTrue(Math.Abs(1.7724538509055159 - (double)actual) < XLHelper.Epsilon);

            actual = XLWorkbook.EvaluateExpr("SqrtPi(2)");
            Assert.IsTrue(Math.Abs(2.5066282746310002 - (double)actual) < XLHelper.Epsilon);
        }

        [TestMethod]
        public void SubtotalAverage()
        {
            var actual = XLWorkbook.EvaluateExpr("Subtotal(1,2,3)");
            Assert.AreEqual(2.5, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(1,""A"",3, 2)");
            Assert.AreEqual(2.5, actual);
        }

        [TestMethod]
        public void SubtotalCount()
        {
            var actual = XLWorkbook.EvaluateExpr("Subtotal(2,2,3)");
            Assert.AreEqual(2.0, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(2,""A"",3)");
            Assert.AreEqual(2.0, actual);
        }

        [TestMethod]
        public void SubtotalCountA()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Subtotal(3,2,3)");
            Assert.AreEqual(2.0, actual);

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(3,"""",3)");
            Assert.AreEqual(1.0, actual);
        }

        [TestMethod]
        public void SubtotalMax()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(4,2,3,""A"")");
            Assert.AreEqual(3.0, actual);
        }

        [TestMethod]
        public void SubtotalMin()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(5,2,3,""A"")");
            Assert.AreEqual(2.0, actual);
        }

        [TestMethod]
        public void SubtotalProduct()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(6,2,3,""A"")");
            Assert.AreEqual(6.0, actual);
        }

        [TestMethod]
        public void SubtotalStDev()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(7,2,3,""A"")");
            Assert.IsTrue(Math.Abs(0.70710678118654757 - (double)actual) < XLHelper.Epsilon);
        }

        [TestMethod]
        public void SubtotalStDevP()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(8,2,3,""A"")");
            Assert.AreEqual(0.5, actual);
        }

        [TestMethod]
        public void SubtotalSum()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(9,2,3,""A"")");
            Assert.AreEqual(5.0, actual);
        }

        [TestMethod]
        public void SubtotalVar()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(10,2,3,""A"")");
            Assert.IsTrue(Math.Abs(0.5 - (double)actual) < XLHelper.Epsilon);
        }

        [TestMethod]
        public void SubtotalVarP()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(11,2,3,""A"")");
            Assert.AreEqual(0.25, actual);
        }

        [TestMethod]
        public void SumSq()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"SumSq(3,4)");
            Assert.AreEqual(25.0, actual);
        }

        [TestMethod]
        public void Asc()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
            Assert.AreEqual("Text", actual);
        }

        [TestMethod]
        public void Clean()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr(String.Format(@"Clean(""A{0}B"")", Environment.NewLine));
            Assert.AreEqual("AB", actual);
        }

        [TestMethod]
        public void Dollar()
        {
            var actual = XLWorkbook.EvaluateExpr("Dollar(12345.123)");
            Assert.AreEqual(TestHelper.CurrencySymbol + " 12,345.12", actual);

            actual = XLWorkbook.EvaluateExpr("Dollar(12345.123, 1)");
            Assert.AreEqual(TestHelper.CurrencySymbol  + " 12,345.1", actual);
        }

        [TestMethod]
        public void Exact()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"A\")");
            Assert.AreEqual(true, actual);

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"a\")");
            Assert.AreEqual(false, actual);
        }

        [TestMethod]
        public void Fixed()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
            Assert.AreEqual("12,345.12", actual);

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
            Assert.AreEqual("12,345.1", actual);

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, FALSE)");
            Assert.AreEqual("12345.1", actual);
        }

        [TestMethod]
        public void Sum()
        {
            var cell = new XLWorkbook().AddWorksheet("Sheet1").FirstCell();
            var fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
            fCell.FormulaA1 = "sum(A1:A2)";

            Assert.AreEqual(3.0, fCell.Value);
        }

        [TestMethod]
        public void MMult()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);
            ws.Cell("A3").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A4").SetValue(3).CellRight().SetValue(5);
            
            Object actual;

            ws.Cell("A5").FormulaA1 = "MMult(A1:B2, A3:B4)";
            actual = ws.Cell("A5").Value;

            Assert.AreEqual(16.0, actual);

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.AreEqual(16.0, actual);

            ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:B2, A3:B4))";
            actual = ws.Cell("A7").Value;

            Assert.AreEqual(102.0, actual);
        }

        [TestMethod]
        public void MDetem()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);


            Object actual;

            ws.Cell("A5").FormulaA1 = "MDeterm(A1:B2)";
            actual = ws.Cell("A5").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));
            
            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));

            ws.Cell("A7").FormulaA1 = "Sum(MDeterm(A1:B2))";
            actual = ws.Cell("A7").Value;

            Assert.IsTrue(XLHelper.AreEqual(-2.0, (double)actual));
        }

        [TestMethod]
        public void MInverse()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(1).CellRight().SetValue(2).CellRight().SetValue(1);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(4).CellRight().SetValue(-1);
            ws.Cell("A3").SetValue(0).CellRight().SetValue(2).CellRight().SetValue(0);


            Object actual;

            ws.Cell("A5").FormulaA1 = "MInverse(A1:C3)";
            actual = ws.Cell("A5").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));

            ws.Cell("A7").FormulaA1 = "Sum(MInverse(A1:C3))";
            actual = ws.Cell("A7").Value;

            Assert.IsTrue(XLHelper.AreEqual(0.5, (double)actual));
        }

        [TestMethod]
        public void TextConcat()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1;
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 1;

            ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

            var r = ws.Cell("C1").Value;
            Assert.AreEqual("The total value is: 4", r.ToString());
        }

        [TestMethod]
        public void Formula_from_another_sheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("ws1");
            ws1.FirstCell().SetValue(1).CellRight().SetFormulaA1("A1 + 1");
            var ws2 = wb.AddWorksheet("ws2");
            ws2.FirstCell().SetFormulaA1("ws1!B1 + 1");
            var v = ws2.FirstCell().Value;
            Assert.AreEqual(3.0, v);
        }
    }
}
