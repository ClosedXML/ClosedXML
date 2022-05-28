using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Expressions in the workbook are optimized before running, e.g. expression <c>(1+14)/5</c> is turned into <c>5</c>.
    /// Since optimization is transparent to the user and happens during parsing, test look into internal structures.
    /// </summary>
    [TestFixture]
    public class OptimizerTests
    {
        [Test]
        public void Binary_operations_with_literal_is_optimized()
        {
            var exp = GetOptimizedExpression("1+2");
            Assert.AreEqual(typeof(Expression), exp.GetType());
            Assert.AreEqual(3, exp._token.Value);
        }

        [Test]
        public void Unary_operations_with_literal_is_optimized()
        {
            var exp = GetOptimizedExpression("-(+7)");
            Assert.AreEqual(typeof(Expression), exp.GetType());
            Assert.AreEqual(-7, exp._token.Value);
        }

        [Test]
        public void Function_call_with_only_literal_values_is_optimized()
        {
            var exp = GetOptimizedExpression("=SUM(2-1,15/5-2,1,2*5-3*3)");
            Assert.AreEqual(typeof(Expression), exp.GetType());
            Assert.AreEqual(4, exp._token.Value);
        }

        private Expression GetOptimizedExpression(string formula)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet() as XLWorksheet;
                return ws.CalcEngine.Parse(formula);
            }
        }
    }
}
