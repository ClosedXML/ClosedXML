using ClosedXML.Excel.CalcEngine.Exceptions;
using Irony.Ast;
using Irony.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLParser;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A parser that takes a string and parses it into concrete syntax tree through XLParser and then
    /// to abstract syntax tree that is used to evaluate the formula.
    /// </summary>
    internal class FormulaParser
    {
        private const string defaultFunctionNameSpace = "_xlfn";

        // Names for binary op terms don't have a const names in the grammar
        private static readonly Dictionary<string, BinaryOp> BinaryOpMap = new()
        {
            { "^", BinaryOp.Exp },
            { "*", BinaryOp.Mult },
            { "/", BinaryOp.Div },
            { "+", BinaryOp.Add },
            { "-", BinaryOp.Sub },
            { "&", BinaryOp.Concat},
            { ">", BinaryOp.Gt},
            { "=", BinaryOp.Eq },
            { "<", BinaryOp.Lt },
            { "<>", BinaryOp.Neq },
            { ">=", BinaryOp.Gte },
            { "<=", BinaryOp.Lte },
        };

        private static readonly IDictionary<string, ErrorExpression.ExpressionErrorType> ErrorMap = new Dictionary<string, ErrorExpression.ExpressionErrorType>(StringComparer.OrdinalIgnoreCase)
        {
            ["#REF!"] = ErrorExpression.ExpressionErrorType.CellReference,
            ["#VALUE!"] = ErrorExpression.ExpressionErrorType.CellValue,
            ["#DIV/0!"] = ErrorExpression.ExpressionErrorType.DivisionByZero,
            ["#NAME?"] = ErrorExpression.ExpressionErrorType.NameNotRecognized,
            ["#N/A"] = ErrorExpression.ExpressionErrorType.NoValueAvailable,
            ["#NULL!"] = ErrorExpression.ExpressionErrorType.NullValue,
            ["#NUM!"] = ErrorExpression.ExpressionErrorType.NumberInvalid
        };

        private readonly CalcEngine _engine;
        private readonly CompatibilityFormulaVisitor _compatibilityVisitor;
        private readonly Dictionary<string, FunctionDefinition> _fnTbl; // table with constants and functions (pi, sin, etc)
        private readonly Parser _parser;

        public FormulaParser(CalcEngine engine, Dictionary<string, FunctionDefinition> fnTbl)
        {
            _engine = engine;
            _compatibilityVisitor = new CompatibilityFormulaVisitor(_engine);
            var grammar = GetGrammar();
            _parser = new Parser(grammar);
            _fnTbl = fnTbl;
        }

        public Expression ParseToAst(string formula)
        {
            try
            {
                var tree = _parser.Parse(formula);
                var root = (Expression)tree.Root.AstNode ?? throw new InvalidOperationException("Formula doesn't have AST root.");
                root = root.Accept(null, _compatibilityVisitor);
                return root;
            }
            catch (NullReferenceException ex) when (ex.StackTrace.StartsWith("   at Irony.Ast.AstBuilder.BuildAst(ParseTreeNode parseNode)"))
            {
                throw new InvalidOperationException($"Unable to parse formula '{formula}'. Some Irony grammar term is missing AST configuration.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExcelFormulaGrammar GetGrammar()
        {
            // Keep AST configuration in same order as is the 'SomeTerm.Rule ='  in in ExcelFormulaGrammar for readability.
            var grammar = new ExcelFormulaGrammar();
            grammar.FormulaWithEq.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Formula.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.ArrayFormula.AstConfig.NodeCreator = CreateNotImplementedNode("array formula");

            grammar.MultiRangeFormula.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Union.AstConfig.NodeCreator = CreateUnionNode;
            grammar.intersectop.AstConfig.NodeCreator = DontCreateNode;

            grammar.Constant.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.Number.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.NumberToken.AstConfig.NodeCreator = CreateNumberNode;
            grammar.Error.AstConfig.NodeCreator = CreateErrorNode;
            grammar.ErrorToken.AstConfig.NodeCreator = DontCreateNode;

            // RefErrorToken is marked with NoAstToken
            grammar.RefError.AstConfig.NodeCreator = CreateErrorNode;
            grammar.RefErrorToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.ConstantArray.AstConfig.NodeCreator = CreateNotImplementedNode("constant array");
            grammar.ArrayColumns.AstConfig.NodeCreator = DontCreateNode;
            grammar.ArrayRows.AstConfig.NodeCreator = DontCreateNode;
            grammar.ArrayConstant.AstConfig.NodeCreator = DontCreateNode;

            grammar.FunctionCall.AstConfig.NodeCreator = GetFunctionCallNodeFactory();
            grammar.Argument.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.FunctionName.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelFunction.AstConfig.NodeCreator = DontCreateNode;

            grammar.Arguments.AstConfig.NodeCreator = DontCreateNode;
            grammar.EmptyArgument.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.EmptyArgumentToken.AstConfig.NodeCreator = CreateEmptyArgumentNode;

            grammar.Bool.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.BoolToken.AstConfig.NodeCreator = CreateBoolNode;

            grammar.Text.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.TextToken.AstConfig.NodeCreator = CreateTextNode;

            // TODO: this is placeholder
            grammar.Reference.AstConfig.NodeCreator = ReferenceNode.CreateReferenceNode;
            grammar.Cell.AstConfig.NodeCreator = DontCreateNode;
            grammar.CellToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.NamedRange.AstConfig.NodeCreator = DontCreateNode;
            grammar.NameToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.NamedRangeCombinationToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.VRange.AstConfig.NodeCreator = DontCreateNode;
            grammar.VRangeToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.HRange.AstConfig.NodeCreator = DontCreateNode;
            grammar.HRangeToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.ReferenceFunctionCall.AstConfig.NodeCreator = CreateReferenceFunctionCallNodeFactory();
            grammar.RefFunctionName.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelConditionalRefFunctionToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelRefFunctionToken.AstConfig.NodeCreator = DontCreateNode;

            // Prefix is only used in Reference term together with ReferenceItem. It is taken care of in CreateReferenceFunctionCallNode.
            grammar.Prefix.AstConfig.NodeCreator = PrefixNode.CreatePrefixNode;
            grammar.SheetToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.SheetQuotedToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.MultipleSheetsToken.AstConfig.NodeCreator = DontCreateNode;

            // DDE formula parsing in XLParser seems to be buggy. It can't parse few examples I have found.
            grammar.DynamicDataExchange.AstConfig.NodeCreator = CreateNotImplementedNode("dynamic data exchange");
            grammar.SingleQuotedStringToken.AstConfig.NodeCreator = DontCreateNode;

            // File is only used in Reference and not directly, so don't use NotImplementedNode since it is never evaluated.
            grammar.File.AstConfig.NodeCreator = FileNode.CreateFileNode;
            grammar.File.SetFlag(TermFlags.AstDelayChildren);

            grammar.UDFunctionCall.AstConfig.NodeCreator = CreateUDFunctionNode;
            grammar.UDFName.AstConfig.NodeCreator = DontCreateNode;
            grammar.UDFToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.StructuredReference.AstConfig.NodeCreator = StructuredReferenceNode.CreateStructuredReferenceNode;
            grammar.StructuredReference.SetFlag(TermFlags.AstDelayChildren);

            // Irony has a few bugs. If it throws a NRE in BuildAst(parseNode), some node is missing a setting to create node for the term.
            grammar.LanguageFlags |= LanguageFlags.CreateAst;
            return grammar;
        }

        private void DontCreateNode(AstContext context, ParseTreeNode parseNode)
        {
            // Don't create an AST node for the parseNode. Its children will use their AstConfig to create their AST nodes.
        }

        private void CreateNumberNode(AstContext context, ParseTreeNode parseNode)
        {
            var value = parseNode.Token.Value is int intValue ? (double)intValue : (double)parseNode.Token.Value;
            parseNode.AstNode = new ScalarNode(value);
        }

        private void CreateBoolNode(AstContext context, ParseTreeNode parseNode)
        {
            var boolValue = string.Equals(parseNode.Token.Text, "TRUE", StringComparison.OrdinalIgnoreCase);
            parseNode.AstNode = new ScalarNode(boolValue);
        }

        private void CreateTextNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new ScalarNode(parseNode.Token.ValueString);
        }

        private void CreateErrorNode(AstContext context, ParseTreeNode parseNode)
        {
            var errorType = ErrorMap[parseNode.ChildNodes.Single().Token.Text];
            parseNode.AstNode = new ErrorExpression(errorType);
        }

        private AstNodeFactory GetFunctionCallNodeFactory()
        {
            return new()
            {
                {
                    For(new [] { "-", "+", "@" }, GrammarNames.Formula),
                    node => new UnaryExpression(node.ChildNodes[0].Term.Name, (Expression)node.ChildNodes[1].AstNode)
                },
                {
                    For(GrammarNames.Formula, "%"),
                    node => new UnaryExpression("%", (Expression)node.ChildNodes[0].AstNode)
                },
                {
                    For(GrammarNames.FunctionName, GrammarNames.Arguments),
                    node => CreateExcelFunctionCallExpression(node.ChildNodes[0], node.ChildNodes[1])
                },
                {
                    For(GrammarNames.Formula, BinaryOpMap.Keys.ToArray(), GrammarNames.Formula),
                    node => new BinaryExpression(BinaryOpMap[node.ChildNodes[1].Term.Name], (Expression)node.ChildNodes[0].AstNode, (Expression)node.ChildNodes[2].AstNode)
                }
            };
        }

        // AST node created by this factory is mostly just copied upwards in the ReferenceNode factory.

        private AstNodeFactory CreateReferenceFunctionCallNodeFactory()
        {
            return new()
            {
                {
                    For(GrammarNames.Reference, ":", GrammarNames.Reference),
                    node => new BinaryExpression(BinaryOp.Range, (Expression)node.ChildNodes[0].AstNode, (Expression)node.ChildNodes[2].AstNode)
                },
                {
                    For(GrammarNames.Reference, GrammarNames.TokenIntersect, GrammarNames.Reference),
                    node => new BinaryExpression(BinaryOp.Intersection, (Expression)node.ChildNodes[0].AstNode, (Expression)node.ChildNodes[2].AstNode)
                },
                {
                    For(GrammarNames.Union),
                    node => (Expression)node.ChildNodes.Single().AstNode
                },
                {
                    For(GrammarNames.RefFunctionName, GrammarNames.Arguments),
                    node => CreateExcelFunctionCallExpression(node.ChildNodes[0], node.ChildNodes[1])
                },
                {
                    For(GrammarNames.Reference, "#"),
                    node => new UnaryExpression("#", (Expression)node.ChildNodes[0].AstNode)
                },
                {
                    For(),
                    node => null
                },
            };
        }

        private void CreateUDFunctionNode(AstContext context, ParseTreeNode parseNode)
        {
            var functionName = parseNode.ChildNodes[0].ChildNodes.Single().Token.Text.WithoutLast(1);

            if (functionName.StartsWith($"{defaultFunctionNameSpace}."))
            {
                parseNode.AstNode = CreateExcelFunctionCallExpression(parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                return;
            }

            var udfFunction = new FunctionDefinition(-1, -1, p => throw new NotImplementedException("Evaluation of custom functions is not implemented."));
            var arguments = parseNode.ChildNodes[1].ChildNodes.Select(treeNode => treeNode.AstNode).Cast<Expression>().ToList();
            parseNode.AstNode = new FunctionExpression(udfFunction, arguments); ;
        }

        private FunctionExpression CreateExcelFunctionCallExpression(ParseTreeNode nameNode, ParseTreeNode argumentsNode)
        {
            var functionName = nameNode.ChildNodes.Single().Token.Text.WithoutLast(1);
            var foundFunction = _fnTbl.TryGetValue(functionName, out FunctionDefinition functionDefinition);
            if (!foundFunction && functionName.StartsWith($"{defaultFunctionNameSpace}."))
                foundFunction = _fnTbl.TryGetValue(functionName.Substring(defaultFunctionNameSpace.Length + 1), out functionDefinition);

            if (!foundFunction)
                throw new NameNotRecognizedException($"The function `{functionName}` was not recognised.");

            var arguments = argumentsNode.ChildNodes.Select(treeNode => treeNode.AstNode).Cast<Expression>().ToList();
            if (functionDefinition.ParmMin != -1 && arguments.Count < functionDefinition.ParmMin)
                throw new ExpressionParseException($"Too few parameters for function '{functionName}'. Expected a minimum of {functionDefinition.ParmMin} and a maximum of {functionDefinition.ParmMax}.");

            if (functionDefinition.ParmMax != -1 && arguments.Count > functionDefinition.ParmMax)
                throw new ExpressionParseException($"Too many parameters for function '{functionName}'.Expected a minimum of {functionDefinition.ParmMin} and a maximum of {functionDefinition.ParmMax}.");

            return new FunctionExpression(functionDefinition, arguments);
        }


        private static AstNodeCreator CreateCopyNode(int childIndex)
        {
            return (context, parseNode) =>
            {
                var copyNode = parseNode.ChildNodes[childIndex];
                parseNode.AstNode = copyNode.AstNode;
            };
        }

        private static AstNodeCreator CreateNotImplementedNode(string featureText)
        {
            return (_, parseNode) => parseNode.AstNode = new NotSupportedNode(featureText);
        }

        private void CreateUnionNode(AstContext context, ParseTreeNode parseNode)
        {
            var unionRangeNode = (Expression)parseNode.ChildNodes[0].AstNode;
            foreach (var referenceNode in parseNode.ChildNodes.Skip(1))
                unionRangeNode = new BinaryExpression(BinaryOp.Union, unionRangeNode, (Expression)referenceNode.AstNode);
            parseNode.AstNode = unionRangeNode;
        }

        #region Old parser compatibility methods

        private void CreateEmptyArgumentNode(AstContext context, ParseTreeNode parseNode)
        {
            // TODO: This is useless for AST, but kept for compatibility reasons with old parser and some function that use it.
            parseNode.AstNode = new EmptyValueExpression();
        }

        #endregion

        private static bool HasMatchingChildren(ParseTreeNode node, params string[] termNames)
        {
            return node.ChildNodes.Select(c => c.Term.Name).SequenceEqual(termNames);
        }

        private class AstNodeFactory : System.Collections.IEnumerable
        {
            private readonly List<KeyValuePair<Condition[], Func<ParseTreeNode, ExpressionBase>>> _factories = new List<KeyValuePair<Condition[], Func<ParseTreeNode, ExpressionBase>>>();

            public void Add(Condition[] cstNodeConditions, Func<ParseTreeNode, ExpressionBase> astNodeFactory)
                => _factories.Add(new KeyValuePair<Condition[], Func<ParseTreeNode, ExpressionBase>>(cstNodeConditions, astNodeFactory));

            public System.Collections.IEnumerator GetEnumerator() => throw new NotSupportedException();

            public static implicit operator AstNodeCreator(AstNodeFactory factory) => factory.CreateNode;

            private void CreateNode(AstContext context, ParseTreeNode parseNode)
            {
                foreach (var factory in _factories)
                {
                    var conditions = factory.Key;
                    var conditionsSatisfied = parseNode.ChildNodes.Count == conditions.Length
                        && parseNode.ChildNodes.Zip(conditions, (n, c) => c.Func(n)).All(x => x);
                    if (conditionsSatisfied)
                    {
                        parseNode.AstNode = factory.Value(parseNode);
                        return;
                    }
                }

                throw new InvalidOperationException($"Failed to convert CST to AST for term {parseNode.Term.Name}.");
            }
        }

        private class Condition
        {
            public Condition(Func<ParseTreeNode, bool> func) => Func = func;

            public Func<ParseTreeNode, bool> Func { get; }

            public static implicit operator Condition(string termName) => new Condition(x => x.Term.Name == termName);
            public static implicit operator Condition(string[] termNames) => new Condition(x => termNames.Contains(x.Term.Name));
            public static implicit operator Condition(Type astNodeType) => new Condition(x => x.AstNode?.GetType() == astNodeType);
        }

        private static Condition[] For(params Condition[] conditions) => conditions;
    }
}
