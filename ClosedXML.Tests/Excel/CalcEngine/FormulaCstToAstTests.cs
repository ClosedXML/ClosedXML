using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using Irony.Parsing;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using static XLParser.GrammarNames;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Tests checking conversion from concrete syntax tree produced by XLParser to abstract syntax tree used by CalcEngine.
    /// Only shape of CST and AST is checked. This is protection againts changes of the grammar and verification that AST if correctly created from CST.
    /// </summary>
    [TestFixture]
    public class FormulaCstToAstTests
    {
        [Test]
        [TestCaseSource(nameof(FormulaWithCstAndAst))]
        public void FormulaProducesCorrectCstAndAst(string formula, string[] expectedCst, Type[] expectedAst)
        {
            var parser = new FormulaParser(new Dictionary<string, FunctionDefinition>());

            var cst = parser.Parse(formula);
            var linearizedCst = LinearizeCst(cst);
            CollectionAssert.AreEqual(expectedCst, linearizedCst);

            var ast = (ExpressionBase)cst.Root.AstNode;
            var linearizedAst = LinearizeAst(ast);
            CollectionAssert.AreEqual(expectedAst, linearizedAst);
        }

        private static System.Collections.IEnumerable FormulaWithCstAndAst()
        {
            // Trees are serialized using standard tree linearization algorithm
            // non-null value - create a new child of current node and move to the child
            // null - go to parent of current node
            // null at the end of traversal are omitted

            // Keep order of test cases same as the order of tested rules ExcelFormulaGrammar. Complex ad hoc formulas should go to the end.

            // Start.Rule = FormulaWithEq
            yield return new TestCaseData(
                "=1",
                new[] { FormulaWithEq, "=", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Start.Rule = Formula
            yield return new TestCaseData(
                "1",
                new[] { Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            /*            
            // Start.Rule = ArrayFormula
            yield return new TestCaseData(
                "{=1}",
                new[] { ArrayFormula, "=", null, Formula, Constant, Number, TokenNumber },
                TODO);
            */

            // Start.Rule = MultiRangeFormula
            yield return new TestCaseData(
                "=A1,B5",
                new[] { MultiRangeFormula, "=", null, Union, Reference, Cell, TokenCell, null, null, null, Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // TODO: Rest of rules

            // -------------- Complex ad hoc test cases --------------
            yield return new TestCaseData(
                "=1+2",
                new[] { FormulaWithEq, "=", null, Formula, FunctionCall, Formula, Constant, Number, TokenNumber, null, null, null, null, "+", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(BinaryExpression), typeof(ScalarNode), null, typeof(ScalarNode) });
        }

        private static LinkedList<string> LinearizeCst(ParseTree tree)
        {
            var result = new LinkedList<string>();
            LinearizeCstNode(tree.Root, result);
            RemoveNullsAtEnd(result);
            return result;

            static void LinearizeCstNode(ParseTreeNode node, LinkedList<string> linearized)
            {
                linearized.AddLast(node.Term.Name);
                foreach (var child in node.ChildNodes)
                    LinearizeCstNode(child, linearized);
                linearized.AddLast((string)null);
            }
        }

        private static readonly LinearizeVisitor _linearizeAstVisitor = new();

        private static LinkedList<Type> LinearizeAst(ExpressionBase root)
        {
            var result = new LinkedList<Type>();
            root.Accept(result, _linearizeAstVisitor);
            RemoveNullsAtEnd(result);
            return result;
        }

        private static void RemoveNullsAtEnd<T>(LinkedList<T> list)
        {
            while (list.Count > 0 && list.Last.Value is null)
                list.RemoveLast();
        }

        private class LinearizeVisitor : DefaultFormulaVisitor<LinkedList<Type>>
        {
            public override ExpressionBase Visit(LinkedList<Type> context, ScalarNode node)
                => LinearizeNode(context, typeof(ScalarNode), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, UnaryExpression node)
                => LinearizeNode(context, typeof(UnaryExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, BinaryExpression node)
                => LinearizeNode(context, typeof(BinaryExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, FunctionExpression node)
                => LinearizeNode(context, typeof(FunctionExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, XObjectExpression node)
                => LinearizeNode(context, typeof(XObjectExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, EmptyValueExpression node)
                => LinearizeNode(context, typeof(EmptyValueExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, ErrorExpression node)
                => LinearizeNode(context, typeof(ErrorExpression), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, NotSupportedNode node)
                => LinearizeNode(context, typeof(NotSupportedNode), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, ReferenceNode node)
                => LinearizeNode(context, typeof(ReferenceNode), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, StructuredReferenceNode node)
                => LinearizeNode(context, typeof(StructuredReferenceNode), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, PrefixNode node)
                => LinearizeNode(context, typeof(PrefixNode), () => base.Visit(context, node));

            public override ExpressionBase Visit(LinkedList<Type> context, FileNode node)
                => LinearizeNode(context, typeof(FileNode), () => base.Visit(context, node));

            private ExpressionBase LinearizeNode(LinkedList<Type> context, Type nodeType, Func<ExpressionBase> func)
            {
                context.AddLast(nodeType);
                var result = func();
                context.AddLast((Type)null);
                return result;
            }
        }
    }
}
