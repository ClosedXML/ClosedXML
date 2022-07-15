using ClosedXML.Excel.CalcEngine;
using Irony.Parsing;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using XLParser;
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
            var dummyFunctions = new Dictionary<string, FunctionDefinition>()
            {
                { "SUM", new FunctionDefinition(0, 255, x => null) },
                { "SIN", new FunctionDefinition(1, 1, x => null) },
                { "RAND", new FunctionDefinition(0, 0, x => null) },
                { "IF", new FunctionDefinition(0, 3, x => null) },
                { "INDEX", new FunctionDefinition(1, 3, x => null) },
            };
            var parser = new FormulaParser(dummyFunctions);

            var cst = parser.Parse(formula);
            var linearizedCst = LinearizeCst(cst);
            CollectionAssert.AreEqual(expectedCst, linearizedCst);

            var ast = (AstNode)cst.Root.AstNode;
            var linearizedAst = LinearizeAst(ast);
            CollectionAssert.AreEqual(expectedAst, linearizedAst);
        }

        private static System.Collections.IEnumerable FormulaWithCstAndAst()
        {
            // Trees are serialized using standard tree linearization algorithm
            // non-null value - create a new child of current node and move to the child
            // null - go to parent of current node
            // nulls at the end of traversal are omitted

            // Keep order of test cases same as the order of tested rules in the ExcelFormulaGrammar. Complex ad hoc formulas should go to the end.
            // A lot of test seem like duplicates, but keep them - goal is to have at least one test for each rule .
            // During XLparser update, compare original grammar with new one and update these tests according to changes.

            // Test are in sync with XLParser 1.5.2

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

            // Start.Rule = ArrayFormula
            yield return new TestCaseData(
                "{=1}",
                new[] { ArrayFormula, "=", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(NotSupportedNode) });

            // Start.Rule = MultiRangeFormula
            yield return new TestCaseData(
                "=A1,B5",
                new[] { MultiRangeFormula, "=", null, Union, Reference, Cell, TokenCell, null, null, null, Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // ArrayFormula.Rule = OpenCurlyParen + eqop + Formula + CloseCurlyParen;
            yield return new TestCaseData(
                "{=1}",
                new[] { ArrayFormula, "=", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(NotSupportedNode) });

            // MultiRangeFormula.Rule = eqop + Union;
            yield return new TestCaseData(
                "=FirstRange,A1B1",
                new[] { MultiRangeFormula, "=", null, Union, Reference, NamedRange, TokenName, null, null, null, Reference, NamedRange, TokenNamedRangeCombination },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // FormulaWithEq.Rule = eqop + Formula;
            yield return new TestCaseData(
                "=1",
                new[] { FormulaWithEq, "=", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Formula.Rule = Reference
            yield return new TestCaseData(
                "A1",
                new[] { Formula, Reference, Cell, TokenCell },
                new[] { typeof(ReferenceNode) });

            // Formula.Rule = Constant
            yield return new TestCaseData(
                "1",
                new[] { Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Formula.Rule = FunctionCall
            yield return new TestCaseData(
                "+1",
                new[] { Formula, FunctionCall, "+", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(UnaryExpression), typeof(ScalarNode) });

            // Formula.Rule = ConstantArray
            yield return new TestCaseData(
                "{1}",
                new[] { Formula, ConstantArray, ArrayColumns, ArrayRows, ArrayConstant, Constant, Number, TokenNumber },
                new[] { typeof(NotSupportedNode) });

            // Formula.Rule = OpenParen + Formula + CloseParen
            yield return new TestCaseData(
                "(1)",
                new[] { Formula, /* ")" is transient */ Formula /* ")" is transient */, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Formula.Rule = ReservedName
            yield return new TestCaseData(
                "_xlnm.SomeName",
                new[] { Formula, ReservedName, TokenReservedName },
                new[] { typeof(NotSupportedNode) });

            // ReservedName.Rule = ReservedNameToken
            yield return new TestCaseData(
                "_xlnm.OtherName",
                new[] { Formula, ReservedName, TokenReservedName },
                new[] { typeof(NotSupportedNode) });

            // Constant.Rule =  Number
            yield return new TestCaseData(
                "1",
                new[] { Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Constant.Rule = Text
            yield return new TestCaseData(
                "\"\"",
                new[] { Formula, Constant, GrammarNames.Text, TokenText },
                new[] { typeof(ScalarNode) });

            // Constant.Rule = Bool
            yield return new TestCaseData(
                "TRUE",
                new[] { Formula, Constant, Bool, TokenBool },
                new[] { typeof(ScalarNode) });

            // Constant.Rule = Error
            yield return new TestCaseData(
                "#DIV/0!",
                new[] { Formula, Constant, Error, TokenError },
                new[] { typeof(ErrorExpression) });

            // Text.Rule = TextToken;
            yield return new TestCaseData(
                "\"Some text with \"\"enclosed\"\" quotes \"",
                new[] { Formula, Constant, GrammarNames.Text, TokenText },
                new[] { typeof(ScalarNode) });

            // Number.Rule = NumberToken;
            yield return new TestCaseData(
                "123.4e-1",
                new[] { Formula, Constant, Number, TokenNumber },
                new[] { typeof(ScalarNode) });

            // Bool.Rule = BoolToken;
            yield return new TestCaseData(
                "TRUE",
                new[] { Formula, Constant, Bool, TokenBool },
                new[] { typeof(ScalarNode) });

            // Error.Rule = ErrorToken;
            yield return new TestCaseData(
                "#VALUE!",
                new[] { Formula, Constant, Error, TokenError },
                new[] { typeof(ErrorExpression) });

            // RefError.Rule = RefErrorToken;
            yield return new TestCaseData(
                "#REF!",
                new[] { Formula, Reference, RefError, TokenRefError },
                new[] { typeof(ErrorExpression) });

            // FunctionCall.Rule = FunctionName + Arguments + CloseParen
            yield return new TestCaseData(
                "SUM(1)",
                new[] { Formula, FunctionCall, FunctionName, ExcelFunction, null, null, Arguments, Argument, Formula, Constant, Number, TokenNumber },
                new[] { typeof(FunctionExpression), typeof(ScalarNode) });

            // FunctionCall.Rule = PrefixOp + Formula
            yield return new TestCaseData(
                "-1",
                new[] { Formula, FunctionCall, "-", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(UnaryExpression), typeof(ScalarNode) });

            // FunctionCall.Rule = Formula + PostfixOp
            yield return new TestCaseData(
                "1%",
                new[] { Formula, FunctionCall, Formula, Constant, Number, TokenNumber, null, null, null, null, "%" },
                new[] { typeof(UnaryExpression), typeof(ScalarNode) });

            // FunctionCall.Rule = Formula + InfixOp + Formula
            yield return new TestCaseData(
                "1+2",
                new[] { Formula, FunctionCall, Formula, Constant, Number, TokenNumber, null, null, null, null, "+", null, Formula, Constant, Number, TokenNumber },
                new[] { typeof(BinaryExpression), typeof(ScalarNode), null, typeof(ScalarNode) });

            // FunctionName.Rule = ExcelFunction;
            yield return new TestCaseData(
                "RAND()",
                new[] { Formula, FunctionCall, FunctionName, ExcelFunction, null, null, Arguments },
                new[] { typeof(FunctionExpression) });

            // Arguments.Rule = MakeStarRule(Arguments, comma, Argument);
            yield return new TestCaseData(
                "SUM(\"1\", TRUE)",
                new[] { Formula, FunctionCall, FunctionName, ExcelFunction, null, null, Arguments,
                    Argument, Formula, Constant, GrammarNames.Text, TokenText, null, null, null, null, null,
                    Argument, Formula, Constant, Bool, TokenBool },
                new[] { typeof(FunctionExpression), typeof(ScalarNode), null, typeof(ScalarNode) });

            // EmptyArgument.Rule = EmptyArgumentToken;
            yield return new TestCaseData(
                "SUM(,)",
                new[] { Formula, FunctionCall, FunctionName, ExcelFunction, null, null, Arguments,
                    Argument, EmptyArgument, TokenEmptyArgument, null, null, null,
                    Argument, EmptyArgument, TokenEmptyArgument },
                new[] { typeof(FunctionExpression), typeof(EmptyValueExpression), null, typeof(EmptyValueExpression) });

            // Argument.Rule = Formula | EmptyArgument;
            yield return new TestCaseData(
                "IF(,1,)",
                new[] { Formula, Reference, ReferenceFunctionCall, RefFunctionName, TokenExcelConditionalRefFunction, null, null, Arguments,
                    Argument, EmptyArgument, TokenEmptyArgument, null, null, null,
                    Argument, Formula, Constant, Number, TokenNumber , null, null, null, null, null,
                    Argument, EmptyArgument, TokenEmptyArgument },
                new[] { typeof(FunctionExpression), typeof(EmptyValueExpression), null, typeof(ScalarNode), null, typeof(EmptyValueExpression) });

            // PrefixOp.Rule = ImplyPrecedenceHere(Precedence.UnaryPreFix) + plusop | ImplyPrecedenceHere(Precedence.UnaryPreFix) + minop | ImplyPrecedenceHere(Precedence.UnaryPreFix) + at;
            yield return new TestCaseData(
                "@A1",
                new[] { Formula, FunctionCall, "@", null, Formula, Reference, Cell, TokenCell },
                new[] { typeof(UnaryExpression), typeof(ReferenceNode) });

            // InfixOp.Rule = expop | mulop | divop | plusop | minop | concatop | gtop | eqop | ltop | neqop | gteop | lteop;
            yield return new TestCaseData(
                "A1^2",
                new[] { Formula, FunctionCall,
                    Formula, Reference, Cell, TokenCell, null, null, null, null,
                    "^", null,
                    Formula, Constant, Number, TokenNumber },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ScalarNode) });

            // PostfixOp.Rule = PreferShiftHere() + percentop;
            yield return new TestCaseData(
                "A1%",
                new[] { Formula, FunctionCall, Formula, Reference, Cell, TokenCell, null, null, null, null, "%" },
                new[] { typeof(UnaryExpression), typeof(ReferenceNode) });

            // Reference.Rule = ReferenceItem
            yield return new TestCaseData(
                "=A1",
                new[] { FormulaWithEq, "=", null, Formula, Reference, Cell, TokenCell },
                new[] { typeof(ReferenceNode) });

            // Reference.Rule = ReferenceFunctionCall
            yield return new TestCaseData(
                "A1:D5",
                new[] { Formula, Reference, ReferenceFunctionCall, Reference, Cell, TokenCell, null, null, null, ":", null,Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // ReferenceFunctionCall.Rule = Reference + intersectop + Reference
            yield return new TestCaseData(
                "A1 D5",
                new[] { Formula, Reference, ReferenceFunctionCall, Reference, Cell, TokenCell, null, null, null, TokenIntersect, null, Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // ReferenceFunctionCall.Rule = OpenParen + Union + CloseParen
            yield return new TestCaseData(
                "(A1,A2)",
                new[] { Formula, Reference, ReferenceFunctionCall, Union, Reference, Cell, TokenCell, null, null, null, Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // XLParser considers the 5 functions that can return reference to be special.
            // ReferenceFunctionCall.Rule = RefFunctionName + Arguments + CloseParen
            yield return new TestCaseData(
                "IF(TRUE, A1, B2)",
                new[] { Formula, Reference, ReferenceFunctionCall, RefFunctionName, TokenExcelConditionalRefFunction, null, null, Arguments,
                    Argument, Formula, Constant, Bool, TokenBool, null, null, null, null, null,
                    Argument, Formula, Reference, Cell, TokenCell, null, null, null, null, null,
                    Argument, Formula, Reference, Cell, TokenCell },
                new[] { typeof(FunctionExpression), typeof(ScalarNode), null, typeof(ReferenceNode), null, typeof(ReferenceNode) });

            // ReferenceFunctionCall.Rule = Reference + hash
            yield return new TestCaseData(
                "A1#",
                new[] { Formula, Reference, ReferenceFunctionCall, Reference, Cell, TokenCell, null, null, null, "#" },
                new[] { typeof(UnaryExpression), typeof(ReferenceNode) });

            // RefFunctionName.Rule = ExcelRefFunctionToken | ExcelConditionalRefFunctionToken;
            yield return new TestCaseData(
                "INDEX(A1,1,1)",
                new[] { Formula, Reference, ReferenceFunctionCall, RefFunctionName, TokenExcelRefFunction, null, null, Arguments,
                    Argument, Formula, Reference, Cell, TokenCell, null, null, null, null, null,
                    Argument, Formula, Constant, Number, TokenNumber, null, null, null, null, null,
                    Argument, Formula, Constant, Number, TokenNumber },
                new[] { typeof(FunctionExpression), typeof(ReferenceNode), null, typeof(ScalarNode), null, typeof(ScalarNode) });

            // Union.Rule = MakePlusRule(Union, comma, Reference);
            yield return new TestCaseData(
                "(A1,A2,A3)",
                new[] { Formula, Reference, ReferenceFunctionCall, Union,
                    Reference, Cell, TokenCell, null, null, null,
                    Reference, Cell, TokenCell, null, null, null,
                    Reference, Cell, TokenCell },
                new[] { typeof(BinaryExpression), typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode), null, null, typeof(ReferenceNode) });

            // ReferenceItem.Rule = Cell
            yield return new TestCaseData(
                "ZZ256",
                new[] { Formula, Reference, Cell, TokenCell },
                new[] { typeof(ReferenceNode) });

            // ReferenceItem.Rule = NamedRange
            yield return new TestCaseData(
                "SomeRange",
                new[] { Formula, Reference, NamedRange, TokenName },
                new[] { typeof(ReferenceNode) });

            // ReferenceItem.Rule = VRange
            yield return new TestCaseData(
                "A:ZZ",
                new[] { Formula, Reference, VerticalRange, TokenVRange },
                new[] { typeof(ReferenceNode) });

            // ReferenceItem.Rule = HRange
            yield return new TestCaseData(
                "15:40",
                new[] { Formula, Reference, HorizontalRange, TokenHRange },
                new[] { typeof(ReferenceNode) });

            // ReferenceItem.Rule = RefError
            yield return new TestCaseData(
                "#REF!",
                new[] { Formula, Reference, RefError, TokenRefError },
                new[] { typeof(ErrorExpression) });

            // ReferenceItem.Rule = UDFunctionCall
            yield return new TestCaseData(
                "Fun()",
                new[] { Formula, Reference, UDFunctionCall, UDFName, TokenUDF, null, null, Arguments },
                new[] { typeof(FunctionExpression) });

            // ReferenceItem.Rule = StructuredReference
            yield return new TestCaseData(
                "[#All]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // UDFunctionCall.Rule = UDFName + Arguments + CloseParen;
            yield return new TestCaseData(
                "CustomUdfFunction(TRUE)",
                new[] { Formula, Reference, UDFunctionCall, UDFName, TokenUDF, null, null, Arguments, Argument, Formula, Constant, Bool, TokenBool },
                new[] { typeof(FunctionExpression), typeof(ScalarNode) });

            // UDFName.Rule = UDFToken;
            yield return new TestCaseData(
                "_xll.CustomFunc()",
                new[] { Formula, Reference, UDFunctionCall, UDFName, TokenUDF, null, null, Arguments },
                new[] { typeof(FunctionExpression) });

            // VRange.Rule = VRangeToken;
            // BUG in XLParser 1.5.2, it considers A:XFD as A:XF union D (named token)
            // yield return new TestCaseData(
            //     "A:XFD",
            //     new[] { Formula, Reference, ReferenceFunctionCall, VerticalRange, TokenVRange },
            //     new[] { typeof(ReferenceNode) });

            // HRange.Rule = HRangeToken;
            yield return new TestCaseData(
                "1:1048576",
                new[] { Formula, Reference, HorizontalRange, TokenHRange },
                new[] { typeof(ReferenceNode) });

            // Cell.Rule = CellToken;
            yield return new TestCaseData(
                "$XFD$1048576",
                new[] { Formula, Reference, Cell, TokenCell },
                new[] { typeof(ReferenceNode) });

            // File.Rule = FileNameNumericToken
            yield return new TestCaseData(
                "[1]!NamedRange",
                new[] { Formula, Reference, Prefix, File, TokenFileNameNumeric, null, null, "!", null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // File.Rule = FileNameEnclosedInBracketsToken
            yield return new TestCaseData(
                "[file with space.xlsx]!NamedRange",
                new[] { Formula, Reference, Prefix, File, TokenFileNameEnclosedInBrackets, null, null, "!", null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // File.Rule = FilePathToken + FileNameEnclosedInBracketsToken
            yield return new TestCaseData(
                @"C:\temp\[file with space.xlsx]!NamedRange",
                new[] { Formula, Reference, Prefix, File, TokenFilePath, null, TokenFileNameEnclosedInBrackets, null, null, "!", null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // File.Rule = FilePathToken + FileName
            yield return new TestCaseData(
                @"C:\temp\file.xlsx!NamedRange",
                new[] { Formula, Reference, Prefix, File, TokenFilePath, null, TokenFileName, null, null, "!", null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // DDX - Windows only interprocess communication standard that uses a shared memory - that is the future :)
            // DynamicDataExchange.Rule = File + exclamationMark + SingleQuotedStringToken;
            yield return new TestCaseData(
                @"[C:\Program files\Company\program.exe]!'arg0,1'",
                new[] { Formula, Reference, DynamicDataExchange, File, TokenFileNameEnclosedInBrackets, null, null, "!", null, TokenSingleQuotedString },
                new[] { typeof(NotSupportedNode) });

            // NamedRange.Rule = NameToken | NamedRangeCombinationToken;
            yield return new TestCaseData(
                "A1Z5",
                new[] { Formula, Reference, NamedRange, TokenNamedRangeCombination },
                new[] { typeof(ReferenceNode) });

            // Prefix.Rule = SheetToken
            yield return new TestCaseData(
                "Sheet1!A1",
                new[] { Formula, Reference, Prefix, TokenSheet, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode) });

            // Prefix.Rule = QuoteS + SheetQuotedToken
            yield return new TestCaseData(
                "'Name with space'!NamedRange",
                new[] { Formula, Reference, Prefix, "'", null, TokenSheetQuoted, null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode) });

            // Prefix.Rule = File + SheetToken
            yield return new TestCaseData(
                "[1]Sheet!A1",
                new[] { Formula, Reference, Prefix, File, TokenFileNameNumeric, null, null, TokenSheet, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // Prefix.Rule = QuoteS + File + SheetQuotedToken
            yield return new TestCaseData(
                @"'C:\temp\[file.xlsx]Sheet1'!A1",
                new[] { Formula, Reference, Prefix, "'", null, File, TokenFilePath, null, TokenFileNameEnclosedInBrackets, null, null, TokenSheetQuoted, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // Prefix.Rule = File + exclamationMark
            yield return new TestCaseData(
                "[file.xlsx]!NamedRange",
                new[] { Formula, Reference, Prefix, File, TokenFileNameEnclosedInBrackets, null, null, "!", null, null, NamedRange, TokenName },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // Prefix.Rule = MultipleSheetsToken
            yield return new TestCaseData(
                "Jan:Feb!A1",
                new[] { Formula, Reference, Prefix, TokenMultipleSheets, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode) });

            // Prefix.Rule = QuoteS + MultipleSheetsQuotedToken
            yield return new TestCaseData(
                "'Human Resources:Facility Management'!A1",
                new[] { Formula, Reference, Prefix, "'", null, TokenMultipleSheetsQuoted, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode) });

            // Prefix.Rule = File + MultipleSheetsToken
            yield return new TestCaseData(
                "[1]Jan:Dec!A1",
                new[] { Formula, Reference, Prefix, File, TokenFileNameNumeric, null, null, TokenMultipleSheets, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // Prefix.Rule = QuoteS + File + MultipleSheetsQuotedToken
            yield return new TestCaseData(
                "'[7]Human Resources:Facility Management'!A1",
                new[] { Formula, Reference, Prefix, "'", null, File, TokenFileNameNumeric, null, null, TokenMultipleSheetsQuoted, null, null, Cell, TokenCell },
                new[] { typeof(ReferenceNode), typeof(PrefixNode), typeof(FileNode) });

            // Prefix.Rule = RefErrorToken
            yield return new TestCaseData(
                "#REF!",
                new[] { Formula, Reference, RefError, TokenRefError },
                new[] { typeof(ErrorExpression) });

            // StructuredReferenceElement.Rule = OpenSquareParen + SRColumnToken + CloseSquareParen
            // BUG in XLParser 1.5.2, FileNameEnclosedInBracketsToken will always take preference, this can never happen. Square parenthesis are transient
            // yield return new TestCaseData(
            //     "[[ColumnName]]",
            //     new[] {  },
            //     new[] { typeof() });

            // StructuredReferenceElement.Rule = OpenSquareParen + NameToken + CloseSquareParen
            // BUG in XLParser 1.5.2, FileNameEnclosedInBracketsToken will always take preference, this can never happen. Square parenthesis are transient
            // yield return new TestCaseData(
            //     "[[ColumnName]]",
            //     new[] {  },
            //     new[] { typeof() });

            // StructuredReferenceElement.Rule = FileNameEnclosedInBracketsToken
            yield return new TestCaseData(
                "[[Column Name]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceTable.Rule = NameToken;
            yield return new TestCaseData(
                "SomeTable[]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceTable, TokenName },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = StructuredReferenceElement
            yield return new TestCaseData(
                "[[Column Name]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = at + StructuredReferenceElement
            yield return new TestCaseData(
                "[@[Sales Amount]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression, "@", null, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = StructuredReferenceElement + colon + StructuredReferenceElement
            yield return new TestCaseData(
                "[[Sales Person]:[Region]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ":", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = at + StructuredReferenceElement + colon + StructuredReferenceElement
            yield return new TestCaseData(
                "[@[Q1]:[Q4]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    "@", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ":", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = StructuredReferenceElement + comma + StructuredReferenceElement
            yield return new TestCaseData(
                "[[Europe],[Asia]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReferenceExpression.Rule = StructuredReferenceElement + comma + StructuredReferenceElement + colon + StructuredReferenceElement
            yield return new TestCaseData(
                "[[Last Year],[Jan]:[Dec]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ":", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // I have no idea why this term is in the XLGrammar grammar. It limits structural references to three columns....
            // StructuredReferenceExpression.Rule = StructuredReferenceElement + comma + StructuredReferenceElement + comma + StructuredReferenceElement
            yield return new TestCaseData(
                "[[First Column], [Second Column], [Third Column]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // More strangeness
            // StructuredReferenceExpression.Rule = StructuredReferenceElement + comma + StructuredReferenceElement + comma + StructuredReferenceElement + colon + StructuredReferenceElement
            yield return new TestCaseData(
                "[[First Column], [Second Column], [Start Range Column]:[End Range Column]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ":", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReference.Rule = StructuredReferenceElement
            yield return new TestCaseData(
                "[Column]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReference.Rule = OpenSquareParen + StructuredReferenceExpression + CloseSquareParen
            yield return new TestCaseData(
                "[[Column]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceExpression, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReference.Rule = StructuredReferenceTable + StructuredReferenceElement
            yield return new TestCaseData(
                "Sales[Jan]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceTable, TokenName, null, null, StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReference.Rule = StructuredReferenceTable + OpenSquareParen + CloseSquareParen
            yield return new TestCaseData(
                "Sales[]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceTable, TokenName },
                new[] { typeof(StructuredReferenceNode) });

            // StructuredReference.Rule = StructuredReferenceTable + OpenSquareParen + StructuredReferenceExpression + CloseSquareParen
            yield return new TestCaseData(
                "DeptSales[[#Totals],[Sales Amount]:[Commission Amount]]",
                new[] { Formula, Reference, StructuredReference, StructuredReferenceTable, TokenName, null, null, StructuredReferenceExpression,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ",", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets, null, null,
                    ":", null,
                    StructuredReferenceElement, TokenFileNameEnclosedInBrackets },
                new[] { typeof(StructuredReferenceNode) });

            // ConstantArray.Rule = OpenCurlyParen + ArrayColumns + CloseCurlyParen;
            yield return new TestCaseData(
                "{1}",
                new[] { Formula, ConstantArray, ArrayColumns, ArrayRows, ArrayConstant, Constant, Number, TokenNumber },
                new[] { typeof(NotSupportedNode) });

            // ArrayColumns.Rule = MakePlusRule(ArrayColumns, semicolon, ArrayRows);
            yield return new TestCaseData(
                "{1;TRUE;#DIV/0!}",
                new[] { Formula, ConstantArray, ArrayColumns,
                    ArrayRows, ArrayConstant, Constant, Number, TokenNumber, null, null, null, null, null,
                    ArrayRows, ArrayConstant, Constant, Bool, TokenBool, null, null, null, null, null,
                    ArrayRows, ArrayConstant, Constant, Error, TokenError },
                new[] { typeof(NotSupportedNode) });

            // ArrayRows.Rule = MakePlusRule(ArrayRows, comma, ArrayConstant);
            yield return new TestCaseData(
                "{1,TRUE,#DIV/0!}",
                new[] { Formula, ConstantArray, ArrayColumns, ArrayRows,
                    ArrayConstant, Constant, Number, TokenNumber, null, null, null, null,
                    ArrayConstant, Constant, Bool, TokenBool, null, null, null, null,
                    ArrayConstant, Constant, Error, TokenError },
                new[] { typeof(NotSupportedNode) });

            // ArrayConstant.Rule = Constant | PrefixOp + Number | RefError;
            yield return new TestCaseData(
                "{#DIV/0!,-1,#REF!}",
                new[] { Formula, ConstantArray, ArrayColumns, ArrayRows,
                    ArrayConstant, Constant, Error, TokenError, null, null, null, null,
                    ArrayConstant, "-", null, Number, TokenNumber, null, null, null,
                    ArrayConstant, RefError, TokenRefError },
                new[] { typeof(NotSupportedNode) });

            // -------------- Complex ad hoc test cases --------------

            // Function within function
            yield return new TestCaseData(
                "=SUM(SIN(IF(A1,1,2)),3)",
                new[] { FormulaWithEq, "=", null, Formula,
                    FunctionCall /* SUM*/, FunctionName, ExcelFunction, null, null, Arguments,
                        Argument, Formula,
                            FunctionCall /* SIN */, FunctionName, ExcelFunction, null, null, Arguments,
                                Argument, Formula, Reference, ReferenceFunctionCall /* IF*/ , RefFunctionName, TokenExcelConditionalRefFunction, null, null, Arguments,
                                    Argument, Formula, Reference, Cell, TokenCell /* A1*/ , null, null, null, null, null,
                                    Argument, Formula, Constant, Number, TokenNumber /* 1 */, null, null, null, null, null,
                                    Argument, Formula, Constant, Number, TokenNumber /* 2 */, null, null, null, null, null,
                                null, null, null, null, null, null, null, null, null,
                        Argument, Formula, Constant, Number, TokenNumber /* 3 */ },
                new[] { typeof(FunctionExpression), /* SUM */
                            typeof(FunctionExpression), /* SIN */
                                typeof(FunctionExpression), /* IF */
                                    typeof(ReferenceNode), null, /* A1 */
                                    typeof(ScalarNode), null, /* 1 */
                                    typeof(ScalarNode), null, /* 2 */
                                null,
                            null,
                            typeof(ScalarNode) /* 3 */ });

            // Multiply reference area with a number
            yield return new TestCaseData(
                "=A1:B2 * 5",
                new[] { FormulaWithEq, "=", null, Formula,
                    FunctionCall,
                        Formula, Reference, ReferenceFunctionCall,
                            Reference, Cell, TokenCell, null, null, null,
                            ":", null,
                            Reference, Cell, TokenCell, null, null, null, null, null, null,
                        "*", null,
                        Formula, Constant, Number, TokenNumber },
                new[] { typeof(BinaryExpression), typeof(BinaryExpression), typeof(ReferenceNode), null, typeof(ReferenceNode), null, null, typeof(ScalarNode) });
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

        private static LinkedList<Type> LinearizeAst(AstNode root)
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
            public override AstNode Visit(LinkedList<Type> context, ScalarNode node)
                => LinearizeNode(context, typeof(ScalarNode), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, UnaryExpression node)
                => LinearizeNode(context, typeof(UnaryExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, BinaryExpression node)
                => LinearizeNode(context, typeof(BinaryExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, FunctionExpression node)
                => LinearizeNode(context, typeof(FunctionExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, XObjectExpression node)
                => LinearizeNode(context, typeof(XObjectExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, EmptyValueExpression node)
                => LinearizeNode(context, typeof(EmptyValueExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, ErrorExpression node)
                => LinearizeNode(context, typeof(ErrorExpression), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, NotSupportedNode node)
                => LinearizeNode(context, typeof(NotSupportedNode), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, ReferenceNode node)
                => LinearizeNode(context, typeof(ReferenceNode), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, StructuredReferenceNode node)
                => LinearizeNode(context, typeof(StructuredReferenceNode), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, PrefixNode node)
                => LinearizeNode(context, typeof(PrefixNode), () => base.Visit(context, node));

            public override AstNode Visit(LinkedList<Type> context, FileNode node)
                => LinearizeNode(context, typeof(FileNode), () => base.Visit(context, node));

            private AstNode LinearizeNode(LinkedList<Type> context, Type nodeType, Func<AstNode> func)
            {
                context.AddLast(nodeType);
                var result = func();
                context.AddLast((Type)null);
                return result;
            }
        }
    }
}
