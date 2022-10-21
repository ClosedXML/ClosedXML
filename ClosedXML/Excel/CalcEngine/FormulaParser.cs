using ClosedXML.Excel.CalcEngine.Exceptions;
using Irony.Ast;
using Irony.Parsing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLParser;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A parser that takes a string and parses it into concrete syntax tree through XLParser and then
    /// to abstract syntax tree that is used to evaluate the formula.
    /// </summary>
    internal class FormulaParser
    {
        /// <summary>
        /// A prefix is that is used for functions that are present in a version of Excel, but aren't present in older versions.
        /// </summary>
        /// <example>
        /// If you write <c>CONCAT(A1,B1)</c> in Excel 2021 (not present in Excel 2013), it is saved to the worksheet file as
        /// <c>_xlfn.CONCAT(A1,B1)</c>, but the Excel GUI will show only <c>CONCAT(A1,B1)</c>, without the <c>_xlfn</c>.
        /// </example>
        private const string DefaultFunctionNameSpace = "_xlfn";

        // Names for unary/binary op terms don't have a const names in the grammar
        private static readonly Dictionary<string, UnaryOp> PrefixOpMap = new(StringComparer.Ordinal)
        {
            {  "+", UnaryOp.Add },
            {  "-", UnaryOp.Subtract },
            {  "@", UnaryOp.ImplicitIntersection }
        };

        private static readonly Dictionary<string, BinaryOp> BinaryOpMap = new(StringComparer.Ordinal)
        {
            { "^", BinaryOp.Exp },
            { "*", BinaryOp.Mult },
            { "/", BinaryOp.Div },
            { "+", BinaryOp.Add },
            { "-", BinaryOp.Sub },
            { "&", BinaryOp.Concat },
            { ">", BinaryOp.Gt },
            { "=", BinaryOp.Eq },
            { "<", BinaryOp.Lt },
            { "<>", BinaryOp.Neq },
            { ">=", BinaryOp.Gte },
            { "<=", BinaryOp.Lte },
        };

        private static readonly Dictionary<string, XLError> ErrorMap = new(StringComparer.OrdinalIgnoreCase)
        {
            ["#REF!"] = XLError.CellReference,
            ["#VALUE!"] = XLError.IncompatibleValue,
            ["#DIV/0!"] = XLError.DivisionByZero,
            ["#NAME?"] = XLError.NameNotRecognized,
            ["#N/A"] = XLError.NoValueAvailable,
            ["#NULL!"] = XLError.NullValue,
            ["#NUM!"] = XLError.NumberInvalid
        };

        private static readonly Dictionary<string, ReferenceItemType> RangeTermMap = new(StringComparer.Ordinal)
        {
            { GrammarNames.Cell, ReferenceItemType.Cell },
            { GrammarNames.VerticalRange, ReferenceItemType.VRange },
            { GrammarNames.HorizontalRange, ReferenceItemType.HRange }
        };

        private static readonly Func<AstNode, bool> IsErrorNode = static node => node is ScalarNode scalarNode && scalarNode.Value.TryPickError(out _);

        private readonly Parser _parser;
        private readonly FunctionRegistry _fnTbl;

        public FormulaParser(FunctionRegistry functionRegistry)
        {
            _parser = new Parser(GetGrammar());
            _fnTbl = functionRegistry;
        }

        internal ParseTree ParseCst(string formulaText)
        {
            try
            {
                return _parser.Parse(formulaText);
            }
            catch (NullReferenceException ex) when (ex.StackTrace.StartsWith("   at Irony.Ast.AstBuilder.BuildAst(ParseTreeNode parseNode)"))
            {
                throw new InvalidOperationException($"Unable to parse formula '{formulaText}'. Some Irony grammar term is missing AST configuration.");
            }
        }

        /// <summary>
        /// Parse a tree into a CSt that also has AST.
        /// </summary>
        public Formula ConvertToAst(ParseTree cst)
        {
            var astContext = new AstContext(_parser.Language);
            if (cst.HasErrors())
                throw new ExpressionParseException($"Unable to parse formula '{cst.SourceText}':\n" + string.Join("\n", cst.ParserMessages.Select(c => $"Location {c.Location.Line}:{c.Location.Column} - {c.Message}")));

            var astBuilder = new AstBuilder(astContext);
            astBuilder.BuildAst(cst);
            var root = (ValueNode)cst.Root.AstNode ?? throw new InvalidOperationException("Formula doesn't have AST root.");

            var flags = astContext.Values.ContainsKey(FormulaFlags.HasSubtotal)
                ? FormulaFlags.HasSubtotal
                : FormulaFlags.None;
            return new Formula(cst.SourceText, root, flags);
        }

        private ExcelFormulaGrammar GetGrammar()
        {
            var grammar = new ExcelFormulaGrammar();
            grammar.FormulaWithEq.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Formula.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.ArrayFormula.AstConfig.NodeCreator = CreateNotImplementedNode("array formula");
            grammar.ArrayFormula.SetFlag(TermFlags.AstDelayChildren);
            grammar.ReservedName.AstConfig.NodeCreator = CreateNotImplementedNode("reserved name");
            grammar.ReservedName.SetFlag(TermFlags.AstDelayChildren);

            grammar.MultiRangeFormula.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Union.AstConfig.NodeCreator = CreateUnionNode;
            grammar.intersectop.SetFlag(TermFlags.NoAstNode);

            grammar.Constant.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.Number.AstConfig.NodeCreator = CreateNumberNode;
            grammar.Number.SetFlag(TermFlags.AstDelayChildren);
            grammar.Bool.AstConfig.NodeCreator = CreateBoolNode;
            grammar.Bool.SetFlag(TermFlags.AstDelayChildren);
            grammar.Text.AstConfig.NodeCreator = CreateTextNode;
            grammar.Text.SetFlag(TermFlags.AstDelayChildren);
            grammar.Error.AstConfig.NodeCreator = CreateErrorNode;
            grammar.Error.SetFlag(TermFlags.AstDelayChildren);
            grammar.RefError.AstConfig.NodeCreator = CreateErrorNode;
            grammar.RefError.SetFlag(TermFlags.AstDelayChildren);
            grammar.ConstantArray.AstConfig.NodeCreator = CreateNotImplementedNode("constant array");
            grammar.ConstantArray.SetFlag(TermFlags.AstDelayChildren);

            grammar.FunctionCall.AstConfig.NodeCreator = GetFunctionCallNodeFactory();
            grammar.FunctionName.SetFlag(TermFlags.NoAstNode);
            grammar.Arguments.AstConfig.NodeCreator = (_, _) => { }; // Irony shouldn't throw if no factory exist, but it does = use empty factory.
            grammar.Argument.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.EmptyArgument.AstConfig.NodeCreator = CreateEmptyArgumentNode;
            grammar.EmptyArgument.SetFlag(TermFlags.AstDelayChildren);

            grammar.Reference.AstConfig.NodeCreator = CreateReferenceNodeFactory();

            // ReferenceItem term is transient - ReferenceNode will create AST nodes for Cell..HRange.
            grammar.Cell.SetFlag(TermFlags.NoAstNode);
            grammar.NamedRange.SetFlag(TermFlags.NoAstNode);
            grammar.VRange.SetFlag(TermFlags.NoAstNode);
            grammar.HRange.SetFlag(TermFlags.NoAstNode);
            grammar.UDFunctionCall.AstConfig.NodeCreator = CreateUDFunctionNode;
            grammar.UDFName.SetFlag(TermFlags.NoAstNode);
            grammar.StructuredReference.AstConfig.NodeCreator = CreateStructuredReferenceNode;
            grammar.StructuredReference.SetFlag(TermFlags.AstDelayChildren);

            grammar.ReferenceFunctionCall.AstConfig.NodeCreator = CreateReferenceFunctionCallNodeFactory();
            grammar.RefFunctionName.SetFlag(TermFlags.NoAstNode);

            // DDE formula parsing in XLParser seems to be buggy. It can't parse any 'in-the-wild' examples I have found.
            grammar.DynamicDataExchange.AstConfig.NodeCreator = CreateNotImplementedNode("dynamic data exchange");
            grammar.DynamicDataExchange.SetFlag(TermFlags.AstDelayChildren);

            grammar.Prefix.AstConfig.NodeCreator = GetPrefixNodeCreator();
            grammar.SheetToken.SetFlag(TermFlags.NoAstNode);
            grammar.SheetQuotedToken.SetFlag(TermFlags.NoAstNode);
            grammar.MultipleSheetsToken.SetFlag(TermFlags.NoAstNode);
            grammar.MultipleSheetsQuotedToken.SetFlag(TermFlags.NoAstNode);

            grammar.File.AstConfig.NodeCreator = CreateFileNodeFactory();
            grammar.File.SetFlag(TermFlags.AstDelayChildren);

            return grammar;
        }

        private void CreateNumberNode(AstContext context, ParseTreeNode parseNode)
        {
            var value = parseNode.ChildNodes.Single().Token.Value;
            parseNode.AstNode = new ScalarNode(value is int intValue ? (double)intValue : (double)value);
        }

        private void CreateBoolNode(AstContext context, ParseTreeNode parseNode)
        {
            var boolValue = string.Equals(parseNode.ChildNodes.Single().Token.Text, "TRUE", StringComparison.OrdinalIgnoreCase);
            parseNode.AstNode = new ScalarNode(boolValue);
        }

        private void CreateTextNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new ScalarNode(parseNode.ChildNodes.Single().Token.ValueString);
        }

        private void CreateErrorNode(AstContext context, ParseTreeNode parseNode)
        {
            var errorType = ErrorMap[parseNode.ChildNodes.Single().Token.Text];
            parseNode.AstNode = new ScalarNode(errorType);
        }

        private AstNodeFactory GetFunctionCallNodeFactory()
        {
            return new()
            {
                {
                    For(PrefixOpMap.Keys.ToArray(), GrammarNames.Formula),
                    node => new UnaryNode(PrefixOpMap[node.ChildNodes[0].Term.Name], (ValueNode)node.ChildNodes[1].AstNode)
                },
                {
                    For(GrammarNames.Formula, "%"),
                    node => new UnaryNode(UnaryOp.Percentage, (ValueNode)node.ChildNodes[0].AstNode)
                },
                {
                    For(GrammarNames.FunctionName, GrammarNames.Arguments),
                    (node, ctx) => CreateExcelFunctionCallExpression(ctx, node.ChildNodes[0], node.ChildNodes[1])
                },
                {
                    For(GrammarNames.Formula, BinaryOpMap.Keys.ToArray(), GrammarNames.Formula),
                    node => new BinaryNode(BinaryOpMap[node.ChildNodes[1].Term.Name], (ValueNode)node.ChildNodes[0].AstNode, (ValueNode)node.ChildNodes[2].AstNode)
                }
            };
        }

        /// <summary>
        /// Reference AST node is significantly different from CST node. It takes Reference, ReferenceFunctionCall and ReferenceItem terms into a reference value
        /// that represent an area of a workbook (ReferenceNode, StructuredReferenceNode) and operations over these areas (BinaryOperation, UnaryOperation, FunctionExpression).
        /// </summary>
        private AstNodeFactory CreateReferenceNodeFactory()
        {
            return new()
            {
                {
                    // ReferenceItem is transient, so its rules are basically merged with Reference - Cell, NamedRange, VRange, HRange
                    For(new[] { GrammarNames.Cell, GrammarNames.VerticalRange, GrammarNames.HorizontalRange }),
                    node => new ReferenceNode(null, RangeTermMap[node.ChildNodes[0].Term.Name], node.ChildNodes[0].ChildNodes.Single().Token.Text)
                },
                {
                    // Named range can be NameToken or NamedRangeCombinationToken. The combination token is there only to detect names like A1A1.
                    For(GrammarNames.NamedRange),
                    node => new NameNode(null, node.ChildNodes[0].ChildNodes.Single().Token.Text)
                },
                {
                    // ReferenceItem:RefError. #REF! error is not grouped with other errors, but is a part of Reference term.
                    For(IsErrorNode),
                    node => (ScalarNode)node.ChildNodes[0].AstNode
                },
                {
                    // ReferenceItem:UDFunctionCall
                    For(GrammarNames.UDFunctionCall),
                    node =>
                    {
                        var fn = (FunctionNode)node.ChildNodes[0].AstNode;
                        return new FunctionNode(null, fn.Name, fn.Parameters);
                    }
                },
                {
                    // ReferenceItem:StructuredReference. TODO: Copy structured reference once implemented
                    For(GrammarNames.StructuredReference),
                    node => new StructuredReferenceNode(null)
                },
                {
                    // ReferenceFunctionCall - Reference + colon + Reference
                    // ReferenceFunctionCall - Reference + intersectop + Reference
                    // ReferenceFunctionCall - Reference + Union + Reference
                    For(typeof(BinaryNode)),
                    node => (BinaryNode)node.ChildNodes[0].AstNode
                },
                {
                    // ReferenceFunctionCall - RefFunctionName + Arguments + CloseParen
                    For(typeof(FunctionNode)),
                    node => (FunctionNode)node.ChildNodes[0].AstNode
                },
                {
                    // ReferenceFunctionCall - Reference + hash
                    For(typeof(UnaryNode)),
                    node => (UnaryNode)node.ChildNodes[0].AstNode
                },
                {
                    // OpenParen + Reference + CloseParen
                    For(typeof(ReferenceNode)),
                    node => (ReferenceNode)node.ChildNodes[0].AstNode
                },
                {
                    // Prefix + ReferenceItem:Cell|NamedRange|VRange|HRange
                    // Split into two branches, because named range is not actually range, but an alias for a formula
                    For(typeof(PrefixNode), new[] { GrammarNames.Cell, GrammarNames.VerticalRange, GrammarNames.HorizontalRange }),
                    node => new ReferenceNode((PrefixNode)node.ChildNodes[0].AstNode, RangeTermMap[node.ChildNodes[1].Term.Name], node.ChildNodes[1].ChildNodes.Single().Token.Text)
                },
                {
                    // Prefix + ReferenceItem:Cell|NamedRange|VRange|HRange
                    For(typeof(PrefixNode), GrammarNames.NamedRange),
                    node => new NameNode((PrefixNode)node.ChildNodes[0].AstNode, node.ChildNodes[1].ChildNodes.Single().Token.Text)
                },
                {
                    // Prefix + ReferenceItem:RefError
                    For(typeof(PrefixNode), IsErrorNode),
                    node =>
                    {
                        // I think =#REF!#REF! was evaluated to #REF! in Excel 2021.
                        return (ScalarNode)node.ChildNodes[1].AstNode;
                    }
                },
                {
                    // Prefix + ReferenceItem:UDFunctionCall
                    For(typeof(PrefixNode), GrammarNames.UDFunctionCall),
                    node =>
                    {
                        var prefix = (PrefixNode)node.ChildNodes[0].AstNode;
                        var fn = (FunctionNode)node.ChildNodes[1].AstNode;
                        return new FunctionNode(prefix, fn.Name, fn.Parameters);
                    }
                },
                {
                    // Prefix + ReferenceItem:StructuredReference. TODO: Copy structured reference once implemented
                    For(typeof(PrefixNode), GrammarNames.StructuredReference),
                    node => new StructuredReferenceNode(null)
                },
                {
                    For(GrammarNames.DynamicDataExchange),
                    node => new NotSupportedNode("dynamic data exchange")
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
                    node => new BinaryNode(BinaryOp.Range, (ValueNode)node.ChildNodes[0].AstNode, (ValueNode)node.ChildNodes[2].AstNode)
                },
                {
                    For(GrammarNames.Reference, GrammarNames.TokenIntersect, GrammarNames.Reference),
                    node => new BinaryNode(BinaryOp.Intersection, (ValueNode)node.ChildNodes[0].AstNode, (ValueNode)node.ChildNodes[2].AstNode)
                },
                {
                    For(GrammarNames.Union),
                    node => (ValueNode)node.ChildNodes.Single().AstNode
                },
                {
                    For(GrammarNames.RefFunctionName, GrammarNames.Arguments),
                    (node, ctx) => CreateExcelFunctionCallExpression(ctx, node.ChildNodes[0], node.ChildNodes[1])
                },
                {
                    For(GrammarNames.Reference, "#"),
                    node => new UnaryNode(UnaryOp.SpillRange, (ValueNode)node.ChildNodes[0].AstNode)
                }
            };
        }

        private AstNodeFactory GetPrefixNodeCreator()
        {
            return new()
            {
                {
                    For(GrammarNames.TokenSheet),
                    node =>
                    {
                        var sheetName = RemoveExclamationMark(node.ChildNodes[0].Token.Text);
                        return new PrefixNode(null, sheetName, null, null);
                    }
                },
                {
                    For("'", GrammarNames.TokenSheetQuoted),
                    node =>
                    {
                        var quotedSheetName = RemoveExclamationMark("'" + node.ChildNodes[1].Token.Text);
                        return new PrefixNode(null, quotedSheetName.UnescapeSheetName(), null, null);
                    }
                },
                {
                    For(typeof(FileNode), GrammarNames.TokenSheet),
                    node =>
                    {
                        var fileNode = (FileNode)node.ChildNodes[0].AstNode;
                        var sheetName = RemoveExclamationMark(node.ChildNodes[1].Token.Text);
                        return new PrefixNode(fileNode, sheetName, null, null);
                    }
                },
                {
                    For("'", typeof(FileNode), GrammarNames.TokenSheetQuoted),
                    node =>
                    {
                        var fileNode = (FileNode)node.ChildNodes[1].AstNode;
                        var quotedSheetName = RemoveExclamationMark("'" + node.ChildNodes[2].Token.Text);
                        return new PrefixNode(fileNode, quotedSheetName.UnescapeSheetName(), null, null);
                    }
                },
                {
                    For(typeof(FileNode), "!"),
                    node =>
                    {
                        var fileNode = (FileNode)node.ChildNodes[0].AstNode;
                        return new PrefixNode(fileNode, null, null, null);
                    }
                },
                {
                    For(GrammarNames.TokenMultipleSheets),
                    node =>
                    {
                        var normalSheets = RemoveExclamationMark(node.ChildNodes[0].Token.Text).Split(':');
                        return new PrefixNode(null, null, normalSheets[0], normalSheets[1]);
                    }
                },
                {
                    For("'", GrammarNames.TokenMultipleSheetsQuoted),
                    node =>
                    {
                        var quotedSheets = RemoveExclamationMark(("'" + node.ChildNodes[1].Token.Text).UnescapeSheetName()).Split(':');
                        return new PrefixNode(null, null, quotedSheets[0], quotedSheets[1]);
                    }
                },
                {
                    For(typeof(FileNode), GrammarNames.TokenMultipleSheets),
                    node =>
                    {
                        var fileNode = (FileNode)node.ChildNodes[0].AstNode;
                        var normalSheets = RemoveExclamationMark(node.ChildNodes[1].Token.Text).Split(':');
                        return new PrefixNode(fileNode, null, normalSheets[0], normalSheets[1]);
                    }
                },
                {
                    For("'", typeof(FileNode), GrammarNames.TokenMultipleSheetsQuoted),
                    node =>
                    {
                        var fileNode = (FileNode)node.ChildNodes[1].AstNode;
                        var quotedSheets = RemoveExclamationMark(("'" + node.ChildNodes[2].Token.Text).UnescapeSheetName()).Split(':');
                        return new PrefixNode(fileNode, null, quotedSheets[0], quotedSheets[1]);
                    }
                },
                {
                    For(GrammarNames.TokenRefError),
                    node =>
                    {
                        // #REF! is a valid sheet name, Token.ValueString is lower case for some reason.
                        return new PrefixNode(null, RemoveExclamationMark(node.ChildNodes[0].Token.Text), null, null);
                    }
                }
            };
        }

        private AstNodeFactory CreateFileNodeFactory()
        {
            return new()
            {
                {
                    For(GrammarNames.TokenFileNameNumeric),
                    node =>
                    {
                        var numberInBrackets = node.ChildNodes[0].Token.Text;
                        var fileNumericIndex = int.Parse(StripBrackets(numberInBrackets), NumberStyles.None);
                        return new FileNode(fileNumericIndex);
                    }
                },
                {
                    For(GrammarNames.TokenFileNameEnclosedInBrackets),
                    node => new FileNode(node.ChildNodes[0].Token.Text)
                },
                {
                    For(GrammarNames.TokenFilePath, GrammarNames.TokenFileNameEnclosedInBrackets),
                    node =>
                    {
                        var filePath = node.ChildNodes[0].Token.Text;
                        var fileName = node.ChildNodes[1].Token.Text;
                        return new FileNode(System.IO.Path.Combine(filePath, StripBrackets(fileName)));
                    }
                },
                {
                    For(GrammarNames.TokenFilePath, GrammarNames.TokenFileName),
                    node =>
                    {
                        var filePath = node.ChildNodes[0].Token.Text;
                        var fileName = node.ChildNodes[1].Token.Text;
                        return new FileNode(System.IO.Path.Combine(filePath, fileName));
                    }
                }
            };
        }

        private void CreateUDFunctionNode(AstContext context, ParseTreeNode parseNode)
        {
            var functionName = parseNode.ChildNodes[0].ChildNodes.Single().Token.Text.WithoutLast(1);

            if (functionName.StartsWith($"{DefaultFunctionNameSpace}."))
            {
                parseNode.AstNode = CreateExcelFunctionCallExpression(context, parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                return;
            }

            var arguments = parseNode.ChildNodes[1].ChildNodes.Select(treeNode => treeNode.AstNode).Cast<ValueNode>().ToList();
            parseNode.AstNode = new FunctionNode(functionName, arguments); ;
        }

        private FunctionNode CreateExcelFunctionCallExpression(AstContext ctx, ParseTreeNode nameNode, ParseTreeNode argumentsNode)
        {
            var functionName = nameNode.ChildNodes.Single().Token.Text.WithoutLast(1);
            var foundFunction = _fnTbl.TryGetFunc(functionName, out var parmMin, out var parmMax);
            if (!foundFunction && functionName.StartsWith($"{DefaultFunctionNameSpace}."))
            {
                functionName = functionName.Substring(DefaultFunctionNameSpace.Length + 1);
                foundFunction = _fnTbl.TryGetFunc(functionName, out parmMin, out parmMax);
            }

            if (!foundFunction)
                throw new NameNotRecognizedException($"The function `{functionName}` was not recognised.");

            var arguments = argumentsNode.ChildNodes.Select(treeNode => treeNode.AstNode).Cast<ValueNode>().ToList();
            if (parmMin != -1 && arguments.Count < parmMin)
                throw new ExpressionParseException($"Too few parameters for function '{functionName}'. Expected a minimum of {parmMin} and a maximum of {parmMax}.");

            if (parmMax != -1 && arguments.Count > parmMax)
                throw new ExpressionParseException($"Too many parameters for function '{functionName}'.Expected a minimum of {parmMin} and a maximum of {parmMax}.");

            if (string.Equals(functionName, @"SUBTOTAL", StringComparison.OrdinalIgnoreCase))
                ctx.Values[FormulaFlags.HasSubtotal] = true;

            return new FunctionNode(functionName, arguments);
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
            var unionRangeNode = (ValueNode)parseNode.ChildNodes[0].AstNode;
            foreach (var referenceNode in parseNode.ChildNodes.Skip(1))
                unionRangeNode = new BinaryNode(BinaryOp.Union, unionRangeNode, (ValueNode)referenceNode.AstNode);
            parseNode.AstNode = unionRangeNode;
        }

        private void CreateEmptyArgumentNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new ScalarNode(ScalarValue.Blank);
        }

        public void CreateStructuredReferenceNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new StructuredReferenceNode(null);
        }

        private static string RemoveExclamationMark(string sheetName)
        {
            if (!sheetName.EndsWith("!"))
                throw new ArgumentException($"'{sheetName}' doesn't end with !", nameof(sheetName));

            return sheetName.Substring(0, sheetName.Length - 1);
        }

        private string StripBrackets(string fileName)
        {
            if (!fileName.StartsWith("[") || !fileName.EndsWith("]"))
                throw new ArgumentException($"'{fileName}' isn't a text in []", nameof(fileName));

            return fileName.Substring(1, fileName.Length - 2);
        }

        private static NodePredicate[] For(params NodePredicate[] conditions) => conditions;

        private class AstNodeFactory : System.Collections.IEnumerable
        {
            private readonly List<KeyValuePair<NodePredicate[], Func<ParseTreeNode, AstContext, AstNode>>> _factories = new();

            public void Add(NodePredicate[] cstNodeConditions, Func<ParseTreeNode, AstNode> astNodeFactory)
                => _factories.Add(new KeyValuePair<NodePredicate[], Func<ParseTreeNode, AstContext, AstNode>>(cstNodeConditions, (node, _) => astNodeFactory(node)));

            public void Add(NodePredicate[] cstNodeConditions, Func<ParseTreeNode, AstContext, AstNode> astNodeFactory)
                => _factories.Add(new KeyValuePair<NodePredicate[], Func<ParseTreeNode, AstContext, AstNode>>(cstNodeConditions, astNodeFactory));

            public System.Collections.IEnumerator GetEnumerator() => throw new NotSupportedException();

            public static implicit operator AstNodeCreator(AstNodeFactory factory) => factory.CreateNode;

            private void CreateNode(AstContext context, ParseTreeNode parseNode)
            {
                // Sequential conditions are slower than binary switch, but it is readable.
                foreach (var factory in _factories)
                {
                    var conditions = factory.Key;
                    var conditionsSatisfied = parseNode.ChildNodes.Count == conditions.Length
                        && parseNode.ChildNodes.Zip(conditions, (n, c) => c.Func(n)).All(x => x);
                    if (conditionsSatisfied)
                    {
                        parseNode.AstNode = factory.Value(parseNode, context);
                        return;
                    }
                }

                throw new InvalidOperationException($"Failed to convert CST to AST for term {parseNode.Term.Name}.");
            }
        }

        private class NodePredicate
        {
            private NodePredicate(Func<ParseTreeNode, bool> func) => Func = func;

            public Func<ParseTreeNode, bool> Func { get; }

            public static implicit operator NodePredicate(string termName) => new(x => x.Term.Name == termName);
            public static implicit operator NodePredicate(string[] termNames) => new(x => termNames.Contains(x.Term.Name));
            public static implicit operator NodePredicate(Type astNodeType) => new(x => x.AstNode?.GetType() == astNodeType);
            public static implicit operator NodePredicate(Func<AstNode, bool> cond) => new(x => x.AstNode is ValueNode astNode && cond(astNode));
        }
    }
}
