using ClosedXML.Excel;
using NUnit.Framework;
using System;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class LogicalTests
    {
        [Test]
        public void If_2_Params_true()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_2_Params_false()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void If_3_Params_true()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_3_Params_false()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
            Assert.AreEqual("F", actual);
        }

        [Test]
        public void If_Comparing_Against_Empty_String()
        {
            Object actual;
            actual = XLWorkbook.EvaluateExpr(@"if(date(2016, 1, 1) = """", ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XLWorkbook.EvaluateExpr(@"if("""" = date(2016, 1, 1), ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XLWorkbook.EvaluateExpr(@"if("""" = 123, ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XLWorkbook.EvaluateExpr(@"if("""" = """", ""A"",""B"")");
            Assert.AreEqual("A", actual);
        }

        [Test]
        public void If_Case_Insensitivity()
        {
            Object actual;
            actual = XLWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void If_Missing_Second_Value_Then_False()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"IF(FALSE, 1,)");
            Assert.AreEqual(false, actual);
        }

        [TestCase("TRUE", false)]
        [TestCase("FALSE", true)]
        [TestCase("IF(TRUE,,)", true)] // Blank
        [TestCase("0", true)]
        [TestCase("0.1", false)]
        [TestCase("\"true\"", false)]
        [TestCase("\"false\"", true)]
        [TestCase("1/0", XLError.DivisionByZero)]
        public void Not(string valueFormula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr($"NOT({valueFormula})"));
        }
    }
}
