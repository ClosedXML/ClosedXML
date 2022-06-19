using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class LogicalTests
    {
        [Test]
        public void If_2_Params_true()
        {
            var actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_2_Params_false()
        {
            var actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void If_3_Params_true()
        {
            var actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_3_Params_false()
        {
            var actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
            Assert.AreEqual("F", actual);
        }

        [Test]
        public void If_Comparing_Against_Empty_String()
        {
            object actual;
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
            object actual;
            actual = XLWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
            Assert.AreEqual(1, actual);
        }
    }
}
