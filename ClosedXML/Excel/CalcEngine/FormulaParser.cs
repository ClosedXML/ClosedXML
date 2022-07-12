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

        // TODO: Remove later, we only need GetExternalObject method, extract it here.
        private readonly CalcEngine _engine;
        private readonly CompatibilityFormulaVisitor _compatibilityVisitor;
        private readonly Dictionary<string, FunctionDefinition> _fnTbl; // table with constants and functions (pi, sin, etc)
        private Dictionary<BnfTerm, BinaryOp> _binaryOpMap;
        private readonly Parser _parser;

        public FormulaParser(CalcEngine engine, Dictionary<string, FunctionDefinition> fnTbl)
        {
            _engine = engine;
            _compatibilityVisitor = new CompatibilityFormulaVisitor(_engine);
            var grammar = GetGrammar();
            _binaryOpMap = new Dictionary<BnfTerm, BinaryOp> {
                { grammar.expop, BinaryOp.Exp },
                { grammar.mulop, BinaryOp.Mult },
                { grammar.divop, BinaryOp.Div },
                { grammar.plusop, BinaryOp.Add },
                { grammar.minop, BinaryOp.Sub },
                { grammar.concatop, BinaryOp.Concat},
                { grammar.gtop, BinaryOp.Gt},
                { grammar.eqop, BinaryOp.Eq },
                { grammar.ltop, BinaryOp.Lt },
                { grammar.neqop, BinaryOp.Neq },
                { grammar.gteop, BinaryOp.Gte },
                { grammar.lteop, BinaryOp.Lte },
            };
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

            grammar.FunctionCall.AstConfig.NodeCreator = CreateFunctionCallNode;
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

            grammar.ReferenceFunctionCall.AstConfig.NodeCreator = CreateReferenceFunctionCallNode;
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
            var valueString = parseNode.Token.ValueString;
            var boolValue = string.Equals(valueString, "TRUE", StringComparison.OrdinalIgnoreCase);
            parseNode.AstNode = new ScalarNode(boolValue);
        }

        private void CreateTextNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new ScalarNode(parseNode.Token.ValueString);
        }

        private void CreateErrorNode(AstContext context, ParseTreeNode parseNode)
        {
            var errorType = ErrorMap[parseNode.ChildNodes.Single().Token.ValueString];
            parseNode.AstNode = new ErrorExpression(errorType);
        }

        private void CreateFunctionCallNode(AstContext context, ParseTreeNode parseNode)
        {
            if (parseNode.ChildNodes.Count == 2)
            {
                var firstTermName = parseNode.ChildNodes[0].Term.Name;
                var secondTermName = parseNode.ChildNodes[1].Term.Name;
                if ((firstTermName == "-" || firstTermName == "+" || firstTermName == "@") && secondTermName == GrammarNames.Formula)
                {
                    parseNode.AstNode = new UnaryExpression(firstTermName, (Expression)parseNode.ChildNodes[1].AstNode);
                    return;
                }
                else if (firstTermName == GrammarNames.Formula && secondTermName == "%")
                {
                    parseNode.AstNode = new UnaryExpression(secondTermName, (Expression)parseNode.ChildNodes[0].AstNode);
                    return;
                }
                else if (firstTermName == GrammarNames.FunctionName
                    && secondTermName == GrammarNames.Arguments)
                {
                    parseNode.AstNode = CreateExcelFunctionCallExpression(parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                    return;
                }
            }
            else if (parseNode.ChildNodes.Count == 3)
            {
                var middleTerm = parseNode.ChildNodes[1].Term;

                if (_binaryOpMap.TryGetValue(middleTerm, out var infixOp)
                    && parseNode.ChildNodes[0].Term.Name == GrammarNames.Formula
                    && parseNode.ChildNodes[2].Term.Name == GrammarNames.Formula)
                {
                    parseNode.AstNode = new BinaryExpression(infixOp, (Expression)parseNode.ChildNodes[0].AstNode, (Expression)parseNode.ChildNodes[2].AstNode);
                    return;
                }
            }

            throw new ExpressionParseException(parseNode);
        }

        // AST node created by this factory is mostly just copied upwards in the ReferenceNode factory.
        private void CreateReferenceFunctionCallNode(AstContext context, ParseTreeNode parseNode)
        {
            if (HasMatchingChildren(parseNode, GrammarNames.Reference, ":", GrammarNames.Reference))
            {
                parseNode.AstNode = new BinaryExpression(BinaryOp.Range, (Expression)parseNode.ChildNodes[0].AstNode, (Expression)parseNode.ChildNodes[2].AstNode);
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Reference, GrammarNames.TokenIntersect, GrammarNames.Reference))
            {
                parseNode.AstNode = new BinaryExpression(BinaryOp.Intersection, (Expression)parseNode.ChildNodes[0].AstNode, (Expression)parseNode.ChildNodes[2].AstNode);
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Union))
            {
                parseNode.AstNode = parseNode.ChildNodes.Single().AstNode;
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.RefFunctionName, GrammarNames.Arguments))
            {
                parseNode.AstNode = CreateExcelFunctionCallExpression(parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Reference, "#"))
            {
                parseNode.AstNode = new UnaryExpression("#", (Expression)parseNode.ChildNodes[0].AstNode);
                return;
            }

            throw new NotSupportedException();
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
    }
}
