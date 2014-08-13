using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class TextTests
    {
        [Test]
        public void Left_Default()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"")");
            Assert.AreEqual("A", actual);
        }

        [Test]
        public void Left_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"", 2)");
            Assert.AreEqual("AB", actual);
        }

        [Test]
        public void Left_BiggerThanLength()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"", 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Left_Empty()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Right_Default()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"")");
            Assert.AreEqual("C", actual);
        }

        [Test]
        public void Right_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"", 2)");
            Assert.AreEqual("BC", actual);
        }

        [Test]
        public void Right_BiggerThanLength()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"", 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Right_Empty()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right("""")");
            Assert.AreEqual("", actual);
        }


        
        [Test]
        public void Mid_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 2, 2)");
            Assert.AreEqual("BC", actual);
        }

        [Test]
        public void Mid_BiggerThanLength()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 1, 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Mid_StartAfter()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 5, 5)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Mid_Empty()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid("""", 1, 1)");
            Assert.AreEqual("", actual);
        }
    }
}