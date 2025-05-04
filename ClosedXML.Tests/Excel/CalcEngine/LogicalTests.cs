using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class LogicalTests
    {
        [Test]
        public void And_IsLogicalConjunction()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, TRUE, TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("AND({TRUE, TRUE}, TRUE)"), Is.EqualTo(true));

                Assert.That(XLWorkbook.EvaluateExpr("AND(FALSE)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, FALSE)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("AND({TRUE, FALSE})"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, {TRUE, FALSE})"), Is.EqualTo(false));
            });
        }

        [TestCase("A1")]
        [TestCase("A1:A5")]
        [TestCase("(A1:A5,B1:B5)")]
        public void And_NoCollectionValues_Error(string range)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate($"AND({range})"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void And_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
        {
            Assert.Multiple(() =>
            {
                // Blank evaluated to false
                Assert.That(XLWorkbook.EvaluateExpr("AND(IF(TRUE,,))"), Is.EqualTo(false));

                // Number coerced to logical
                Assert.That(XLWorkbook.EvaluateExpr("AND(0)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("AND(0.1)"), Is.EqualTo(true));

                // Text coerced to logical
                Assert.That(XLWorkbook.EvaluateExpr("AND(\"FALSE\")"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("AND(\"TRUE\")"), Is.EqualTo(true));
            });
        }

        [Test]
        public void And_UnconvertableScalarArgumentsSkipped()
        {
            Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE,\"z\")"), Is.EqualTo(true));
        }

        [Test]
        public void And_OnlyLogicalOrNumberElementsOfCollectionUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // 0 is a number and is converted to logical
            ws.Cell("A1").Value = 0;
            Assert.That(ws.Evaluate("AND(TRUE,A1)"), Is.EqualTo(false));

            // false is a logical
            ws.Cell("A2").Value = false;
            Assert.That(ws.Evaluate("AND(TRUE,A2)"), Is.EqualTo(false));

            // Text is not converted and thus skipped for evaluation
            ws.Cell("A3").Value = "FALSE";
            Assert.That(ws.Evaluate("AND(TRUE,A3)"), Is.EqualTo(true));

            ws.Cell("A4").Value = "some text";
            Assert.That(ws.Evaluate("AND(TRUE,A4)"), Is.EqualTo(true));
        }

        [Test]
        public void If_2_Params_true()
        {
            object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
            Assert.That(actual, Is.EqualTo("T"));
        }

        [Test]
        public void If_2_Params_false()
        {
            object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void If_3_Params_true()
        {
            object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
            Assert.That(actual, Is.EqualTo("T"));
        }

        [Test]
        public void If_3_Params_false()
        {
            object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
            Assert.That(actual, Is.EqualTo("F"));
        }

        [Test]
        public void If_Comparing_Against_Empty_String()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr(@"if(date(2016, 1, 1) = """", ""A"",""B"")");
            Assert.That(actual, Is.EqualTo("B"));

            actual = XLWorkbook.EvaluateExpr(@"if("""" = date(2016, 1, 1), ""A"",""B"")");
            Assert.That(actual, Is.EqualTo("B"));

            actual = XLWorkbook.EvaluateExpr(@"if("""" = 123, ""A"",""B"")");
            Assert.That(actual, Is.EqualTo("B"));

            actual = XLWorkbook.EvaluateExpr(@"if("""" = """", ""A"",""B"")");
            Assert.That(actual, Is.EqualTo("A"));
        }

        [Test]
        public void If_Case_Insensitivity()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
            Assert.That(actual, Is.EqualTo(1));
        }

        [Test]
        public void If_CanReturnReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("ISREF(IF(TRUE, A1))"), Is.EqualTo(true));
                Assert.That(ws.Evaluate("ISREF(IF(FALSE,, A1))"), Is.EqualTo(true));
            });
        }

        [Test]
        public void If_has_scalar_condition_and_range_values()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Cell("B1").InsertData(new[] { 4, 5, 6 });
            ws.Cell("C1").InsertData(new[] { true, false, true });
            for (var row = 1; row <= 4; ++row)
                ws.Cell(row, 4).FormulaA1 = "SUM(IF(C1:C3, A1:A3, B1:B3))";

            Assert.Multiple(() =>
            {
                // Condition is implicitely intersected, because it's a scalar parameter
                Assert.That(ws.Cell("D1").Value, Is.EqualTo(6));
                Assert.That(ws.Cell("D2").Value, Is.EqualTo(15));
                Assert.That(ws.Cell("D3").Value, Is.EqualTo(6));
                Assert.That(ws.Cell("D4").Value, Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void If_ConditionError_ReturnError()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"IF(1/0, ""T"", ""F"")"), Is.EqualTo(XLError.DivisionByZero));
        }

        [Test]
        public void If_ConditionCoercedToLogical()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate(@"IF(A1, ""T"", ""F"")"), Is.EqualTo("F"));

                Assert.That(ws.Evaluate(@"IF(""TRUE"", ""T"", ""F"")"), Is.EqualTo("T"));
                Assert.That(ws.Evaluate(@"IF(""FALSE"", ""T"", ""F"")"), Is.EqualTo("F"));
                Assert.That(ws.Evaluate(@"IF(""text"", ""T"", ""F"")"), Is.EqualTo(XLError.IncompatibleValue));

                Assert.That(ws.Evaluate(@"IF(1, ""T"", ""F"")"), Is.EqualTo("T"));
                Assert.That(ws.Evaluate(@"IF(0, ""T"", ""F"")"), Is.EqualTo("F"));
            });
        }

        [Test]
        public void If_MissingValues_ReturnBlank()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr(@"ISBLANK(IF(TRUE,,))"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr(@"ISBLANK(IF(FALSE,,))"), Is.EqualTo(true));
            });
        }

        [Test]
        public void IfError_FirstArgumentNonError_ReturnFirstArgument()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(IF(TRUE,), 5))"), Is.EqualTo(true));

                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(FALSE, 5)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(TRUE, 5)"), Is.EqualTo(true));

                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(0, 5)"), Is.EqualTo(0.0));
                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(-2, 5)"), Is.EqualTo(-2.0));

                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(\"\", 5)"), Is.EqualTo(""));
                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(\"text\", 5)"), Is.EqualTo("text"));
            });
        }

        [Test]
        public void IfError_FirstArgumentError_ReturnSecondArgument()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(1/0, \"text\")"), Is.EqualTo("text"));

                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(#REF!, #NAME?)"), Is.EqualTo(XLError.NameNotRecognized));
                Assert.That(XLWorkbook.EvaluateExpr("IFERROR(#NULL!, TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(#VALUE!,IF(TRUE,)))"), Is.EqualTo(true));
            });
        }

        [Test]
        public void IfError_ReferenceNeverReturned()
        {
            // Unlike IF, IFERROR doesn't return reference
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate("ISREF(IFERROR(#VALUE!, A1))"), Is.EqualTo(false));
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
            Assert.That(XLWorkbook.EvaluateExpr($"NOT({valueFormula})"), Is.EqualTo(expectedResult));
        }

        [Test]
        public void Or_IsLogicalDisjunction()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE, TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE, FALSE, TRUE)"), Is.EqualTo(true));
                Assert.That(XLWorkbook.EvaluateExpr("OR({FALSE, TRUE}, FALSE)"), Is.EqualTo(true));

                Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE, FALSE)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("OR({FALSE, FALSE})"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE, {FALSE, FALSE})"), Is.EqualTo(false));
            });
        }

        [TestCase("A1")]
        [TestCase("A1:A5")]
        [TestCase("(A1:A5,B1:B5)")]
        public void Or_NoCollectionValues_Error(string range)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate($"OR({range})"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Or_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
        {
            Assert.Multiple(() =>
            {
                // Blank evaluated to false
                Assert.That(XLWorkbook.EvaluateExpr("OR(IF(TRUE,,))"), Is.EqualTo(false));

                // Number coerced to logical
                Assert.That(XLWorkbook.EvaluateExpr("OR(0)"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("OR(0.1)"), Is.EqualTo(true));

                // Text coerced to logical
                Assert.That(XLWorkbook.EvaluateExpr("OR(\"FALSE\")"), Is.EqualTo(false));
                Assert.That(XLWorkbook.EvaluateExpr("OR(\"TRUE\")"), Is.EqualTo(true));
            });
        }

        [Test]
        public void Or_UnconvertableScalarArgumentsSkipped()
        {
            Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE,\"z\")"), Is.EqualTo(true));
        }

        [Test]
        public void Or_OnlyLogicalOrNumberElementsOfCollectionUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // 1 is a number and is converted to logical
            ws.Cell("A1").Value = 1;
            Assert.That(ws.Evaluate("OR(FALSE,A1)"), Is.EqualTo(true));

            // false is a logical
            ws.Cell("A2").Value = true;
            Assert.That(ws.Evaluate("OR(FALSE,A2)"), Is.EqualTo(true));

            // Text is not converted and thus skipped for evaluation
            ws.Cell("A3").Value = "TRUE";
            Assert.That(ws.Evaluate("OR(FALSE,A3)"), Is.EqualTo(false));

            ws.Cell("A4").Value = "some text";
            Assert.That(ws.Evaluate("OR(FALSE,A4)"), Is.EqualTo(false));
        }
    }
}
