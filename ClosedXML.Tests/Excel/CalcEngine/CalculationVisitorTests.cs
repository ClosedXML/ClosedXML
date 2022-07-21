using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CalculationVisitorTests
    {
        [TestCase("=COS(0)")]
        public void DevTest(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 10;
            ws.Cell("A3").Value = 100;
            var parser = new FormulaParser(CreateRegistry());
            var cst = parser.Parse(formula);
            var ast = (AstNode)cst.Root.AstNode;

            var context = new CalcContext(null, CultureInfo.InvariantCulture, wb, (XLWorksheet)ws, new XLAddress((XLWorksheet)ws, 2, 5, true, true));
            var func = new FunctionRegistry();
            OperationsTests.TestFuncRegistry.Register(func);

            var visitor = new CalculationVisitor(func);
            var result = ast.Accept(context, visitor);

            if (context.UseImplicitIntersection && result.IsT4)
            {
                result = result.AsT4[0, 0].ToAnyValue();
            }
            Assert.AreEqual(AnyValue.FromT1(new Number1(1)), result);
        }

        [Test]
        public void ScalarNode_ReturnsLogicalValue()
        {
            Assert.Fail();
        }

        private FunctionRegistry CreateRegistry()
        {
            var dummyFunctions = new FunctionRegistry();
            dummyFunctions.RegisterFunction("SUM", 0, 255, x => null);
            dummyFunctions.RegisterFunction("SIN", 1, 1, x => null);
            dummyFunctions.RegisterFunction("RAND", 0, 0, x => null);
            dummyFunctions.RegisterFunction("IF", 0, 3, x => null);
            dummyFunctions.RegisterFunction("INDEX", 1, 3, x => null);
            dummyFunctions.RegisterFunction("COS", 1, 1, x => null);
            return dummyFunctions;
        }

        [Test]
        public void EvaluationWithoutWorksheet()
        {
            var result = XLWorkbook.EvaluateExpr("=1+2");
            Assert.AreEqual(3, result);
        }

        [Test]
        public void EvaluationCanCallFunction()
        {
            var result = XLWorkbook.EvaluateExpr("=COS(0)");
            Assert.AreEqual(1, result);
        }

        [Test]
        public void EvaluationCanCallFunctionWithReference()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Math.PI;

            var result = ws.Evaluate("=SIN(A1)");
            Assert.That(result, Is.EqualTo(0.0).Within(1e-10));
        }

        [Test]
        public void Evaluation_works_with_implicit_intersection()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = -Math.PI/2.0;
            ws.Cell("A3").Value = 3;
            ws.Cell("B2").FormulaA1 = "=ABS(A1:A3)";

            var result = ws.Cell("B2").Value;
            Assert.That(result, Is.EqualTo(Math.PI / 2.0).Within(1e-10));
        }

    }
}
