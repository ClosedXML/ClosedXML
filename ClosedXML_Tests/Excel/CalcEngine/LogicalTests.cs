using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
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
    }
}