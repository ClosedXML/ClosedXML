using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Range>;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CalculationVisitorTests
    {
        private readonly static Dictionary<string, FunctionDefinition> dummyFunctions = new Dictionary<string, FunctionDefinition>()
            {
                { "SUM", new FunctionDefinition(0, 255, x => null) },
                { "SIN", new FunctionDefinition(1, 1, x => null) },
                { "RAND", new FunctionDefinition(0, 0, x => null) },
                { "IF", new FunctionDefinition(0, 3, x => null) },
                { "INDEX", new FunctionDefinition(1, 3, x => null) },
                { "COS", new FunctionDefinition(1, 1, x => null) },
            };

        [TestCase("=COS(0)")]
        public void DevTest(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 10;
            ws.Cell("A3").Value = 100;
            var parser = new FormulaParser(dummyFunctions);
            var cst = parser.Parse(formula);
            var ast = (AstNode)cst.Root.AstNode;

            var context = new CalcContext(CultureInfo.InvariantCulture, (XLWorksheet)ws, new XLAddress((XLWorksheet)ws, 2, 5, true, true));
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
    }
}
