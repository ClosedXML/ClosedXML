using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using static ClosedXML.Excel.CalcEngine.ErrorExpression;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Tests that verify that we can parse formulas and evaluate them. Take a look at XLParser ExcelFormulaGrammar.cs and each rule + its transformation into Abstract Syntax Tree is checked here.
    /// </summary>
    [TestFixture]
    public class FormulaParserTests
    {
        #region Start.Rule

        [TestCase]
        public void Formula_string_can_starting_with_an_equal_sign()
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("=1"));
        }

        [TestCase]
        public void Formula_string_can_omit_starting_equal_sign()
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("1"));
        }

        [TestCase]
        public void Formula_string_can_be_an_array_formula()
        {
            AssertCanParseButNotEvaluate("{=1}", "Evaluation of array formula is not implemented.");
        }

        [TestCase]
        public void Root_formula_string_can_be_union_without_parenthesis()
        {
            // Root of a formula string is pretty much the only place where reference union can be without parenthesis. Elsewhere it must have
            // parthesis to avoid misusing union op (coma) with a separation of arguments in a function call.
            AssertCanParseButNotEvaluate("=A1,A3", "Evaluation of range union operator is not implemented.");
        }

        #endregion

        #region Formula.Rule

        [TestCase]
        public void Formula_can_be_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Text";
            Assert.AreEqual("Text", ws.Evaluate("=A1"));
        }

        [TestCase("=1", 1)]
        [TestCase("=\"text\"", "text")]
        [TestCase("=TRUE", true)]
        public void Formula_can_be_constant(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=SUM(1,2)", 3)]
        [TestCase("=2+3", 5)]
        [TestCase("=-3", -3)]
        // TODO:[TestCase("=150%", 1.5)]
        public void Formula_can_be_function_call(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase]
        public void Formula_can_be_constant_array()
        {
            AssertCanParseButNotEvaluate("={1,2,3;4,5,6}", "Evaluation of constant array is not implemented.");
        }

        [TestCase("=(1)", 1)]
        [TestCase("=(\"text\")", "text")]
        public void Formula_can_be_another_formula_in_parenthesis(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }
        #endregion

        #region Constant.Rule
        [TestCase("=1", 1)]
        [TestCase("=1.5", 1.5)]
        [TestCase("=1.23e2", 123)]
        [TestCase("=1.23e-1", 0.123)]
        [TestCase("=1.23e+3", 1230)]
        public void Constant_can_be_number(string formula, double expectedNumber)
        {
            Assert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=\"text\"", "text")]
        [TestCase("=\"first line\nsecond line\"", "first line\nsecond line")]
        [TestCase("=\"we'll\"", "we'll")]
        [TestCase("=\"use two double quote \"\" to nest quotes\"", "use two double quote \" to nest quotes")]
        public void Constant_can_be_text(string formula, string expectedText)
        {
            Assert.AreEqual(expectedText, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=TRUE", true)]
        [TestCase("=FALSE", false)]
        [TestCase("=tRuE", true)]
        public void Constant_can_be_bool(string formula, bool expectedBool)
        {
            Assert.AreEqual(expectedBool, XLWorkbook.EvaluateExpr(formula));
        }

        // #REF! is coverted by a different rule, so it is not here.
        [TestCase("#VALUE!", ExpressionErrorType.CellValue)]
        [TestCase("#DIV/0!", ExpressionErrorType.DivisionByZero)]
        [TestCase("#NAME?", ExpressionErrorType.NameNotRecognized)]
        [TestCase("#N/A", ExpressionErrorType.NoValueAvailable)]
        [TestCase("#NULL!", ExpressionErrorType.NullValue)]
        [TestCase("#NUM!", ExpressionErrorType.NumberInvalid)]
        public void Constant_can_be_error(string formula, object expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr(formula));
        }
        #endregion

        // Function call from XLParser is anything that takes arguments and uses some transformation (e.g. addition, excel function, unary operation..)
        #region FunctionCall.Rule

        [TestCase("=COS(0)", 1)]
        [TestCase("=SUM(1,2,3)", 6)]
        public void FunctionCall_can_be_excel_predefined_function(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=+1", 1)]
        [TestCase("=-1", -1)]
        //        [TestCase("=@A1", 1)]
        public void FunctionCall_can_be_unary_prefix_operation(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=75%", 0.75)]
        [Ignore("Percent operation not yet implemented.")]
        public void FunctionCall_can_be_unary_postfix_operation(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=2^3", 8)]
        [TestCase("=4^1.5", 8)]
        [TestCase("=3*2", 6)]
        [TestCase("=6/2", 3)]
        [TestCase("=3/2", 1.5)]
        [TestCase("=1+2", 3)]
        [TestCase("=3-5", -2)]
        //        [TestCase(@"=""A"" & ""B""", "AB")]
        [TestCase("=2>1", true)]
        [TestCase("=1>2", false)]
        [TestCase("=5=5", true)]
        [TestCase("=1=2", false)]
        [TestCase("=1<2", true)]
        [TestCase("=2<1", false)]
        [TestCase("=2<>1", true)]
        [TestCase("=3<>3", false)]
        [TestCase("=2>=1", true)]
        [TestCase("=2>=2", true)]
        [TestCase("=1>=2", false)]
        [TestCase("=1<=2", true)]
        [TestCase("=1<=1", true)]
        [TestCase("=2<=1", false)]
        public void FunctionCall_can_be_binary_infix_operation(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }
        #endregion

        #region Argument.Rule

        [TestCase("=PMT(0,1,1000,,1)", -1000)]
        public void Empty_arguments_are_passed_to_function(string formula, object expectedValue)
        {
            Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
        }

        #endregion

        #region Reference.Rule

        [TestCase("=A1", 1)]
        [TestCase("=TestRangeName", 5)]
        //        [TestCase("=UndefinedRangeName", ExpressionErrorType.NameNotRecognized)]
        public void Reference_can_be_reference_item(string formula, object expectedValue)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 5;
            ws.Range("A2:A2").AddToNamed("TestRangeName");

            Assert.AreEqual(expectedValue, ws.Evaluate(formula));
        }

        [TestCase]
        public void Reference_can_be_reference_function_call()
        {
            // XLParser considers a limited subset of predefined functions (IF, CHOOSE, INDEX...) to be different from other predefined function because they can return reference.
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("=IF(FALSE,1,2)"));
        }

        [TestCase]
        public void Reference_can_be_another_reference_in_parenthesis()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;

            Assert.AreEqual(1, ws.Evaluate("=(A1)"));
        }

        [TestCase]
        public void Reference_can_be_reference_item_with_prefix()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("Sheet2");
            ws2.Cell("A1").Value = 1;

            Assert.AreEqual(1, ws1.Evaluate("=Sheet2!  A1"));
        }

        [TestCase]
        [Ignore("XLParser issue #57")]
        public void Reference_can_be_dynamic_data_exchange()
        {
            AssertCanParseButNotEvaluate("=Sdemo123|tik!'id1?req?AAPL_STK_SMART_USD_~/'", "Evaluation of dynamic data exchange is not implemented.");
        }

        #endregion

        #region ReferenceFunctionCall.Rule

        [TestCase]
        public void Reference_function_call_can_be_binary_range_of_two_references()
        {
            AssertCanParseButNotEvaluate("=A1:A3:C2", "Evaluation of binary range operator is not implemented.");
        }

        [TestCase]
        public void Reference_function_call_can_be_intersection_of_two_references()
        {
            AssertCanParseButNotEvaluate("=A1:A3 A2:B2", "Evaluation of range intersection operator is not implemented.");
        }

        [TestCase]
        public void Reference_function_call_can_be_union_in_parenthesis()
        {
            AssertCanParseButNotEvaluate("=(A1:A3,A2:B2,B1:B4)", "Evaluation of range union operator is not implemented.");
        }

        [TestCase]
        public void Reference_function_call_can_be_reference_function()
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("=IF(TRUE,1,2)"));
        }

        [TestCase]
        public void Reference_function_call_can_be_reference_with_spill_range_operator()
        {
            AssertCanParseButNotEvaluate("=A1#", "Evaluation of spill range operator is not implemented.");
        }

        #endregion

        #region RefFunctionName.Rule

        [TestCase("=IF(FALSE,1,2)", 2)]
        // [TestCase("=CHOOSE(2,\"A\",\"B\",73)", "B")] Not implemented
        public void Ref_function_name_can_be_excel_ref_conditional_function(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("=INDEX(A1:B2,1,2)", "Lemons")]
        //[TestCase("=OFFSET(C4,-1,-2)", "Pears")] Not implemented
        //[TestCase("=INDIRECT(\"A2\")", "Bananas")] Not implemented
        public void Ref_function_name_can_be_excel_ref_function(string formula, object expectedValue)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Apples";
            ws.Cell("B1").Value = "Lemons";
            ws.Cell("A2").Value = "Bananas";
            ws.Cell("B2").Value = "Pears";
            Assert.AreEqual(expectedValue, ws.Evaluate(formula));
        }

        #endregion

        #region ReferenceItem.Rule
        // Reference item is transient and is thus inside the reference

        [TestCase]
        public void Reference_item_can_be_cell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;

            Assert.AreEqual(1, ws.Evaluate("=A1"));
        }

        [TestCase("TestRange")]
        [TestCase("A1A1")]
        public void Reference_item_can_be_named_range(string rangeName)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:C4").SetValue(1).AddToNamed(rangeName);

            Assert.AreEqual(12, ws.Evaluate($"=SUM({rangeName})"));
        }

        [TestCase]
        public void Reference_item_can_be_vertical_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:C4").SetValue(1);

            Assert.AreEqual(8, ws.Evaluate("=SUM(A:B)"));
        }

        [TestCase]
        public void Reference_item_can_be_horizontal_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:C4").SetValue(1);

            Assert.AreEqual(3, ws.Evaluate("=SUM(2:2)"));
        }

        [TestCase]
        public void Reference_item_can_be_ref_error()
        {
            Assert.AreEqual(ExpressionErrorType.CellReference, XLWorkbook.EvaluateExpr("#REF!"));
        }

        [TestCase]
        public void Reference_item_can_be_user_defined_function_call()
        {
            AssertCanParseButNotEvaluate("=CustomFunction(1)", "Evaluation of custom functions is not implemented.");
        }

        [TestCase]
        public void Reference_item_can_be_structured_reference()
        {
            AssertCanParseButNotEvaluate("=SomeTable[#Data]", "Evaluation of structured references is not implemented.");
        }

        #endregion

        #region ConstantArray.Rule

        [TestCase("={1}")]
        [TestCase("={1,2,3,4}")]
        [TestCase("={1,2;3,4}")]
        [TestCase("={+1,#REF!,\"Text\";FALSE,#DIV/0!,-1.5}")]
        public void Const_array_can_have_only_scalars(string formula)
        {
            AssertCanParseButNotEvaluate(formula, "Evaluation of constant array is not implemented.");
        }

        #endregion

        #region Prefix.Rule

        // No quotes
        [TestCase("=Sheet5!A1", "Sheet5")]
        [TestCase("=Test_sheet!A1", "Test_sheet")]
        // Sheet with quotes
        [TestCase("='Test Sheet'!A1", "Test Sheet")]
        [TestCase("='Test-Sheet'!A1", "Test-Sheet")]
        [TestCase("='^%>;-+'!A1", "^%>;-+")]
        // Sheet can be named as #REF! error
        [TestCase("=#REF!A1", "#REF")]
        public void Prefix_can_be_sheet_token(string formula, string sheetName)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet(sheetName);
            ws.Cell("A1").Value = 5;
            Assert.AreEqual(5, ws.Evaluate(formula));
        }

        [TestCase("=Sheet1:Sheet5!A1")]
        [TestCase("=Jan:Dec!A1")]
        public void Prefix_can_be_sheets_for_3d_reference(string formula)
        {
            AssertCanParseButNotEvaluate(formula, "Evaluation of reference is not implemented.");
        }

        [TestCase("=[1]Sheet4!A1")]
        [TestCase("=[C:\\file.xlsx]Sheet1!A1")]
        public void Prefix_can_be_file_and_sheet_token(string formula)
        {
            AssertCanParseButNotEvaluate(formula, "Evaluation of reference is not implemented.");
        }

        #endregion

        private static void AssertCanParseButNotEvaluate(string formula, string notSupportedMessage)
        {
            using var wb = new XLWorkbook();
            var calcEngine = new XLCalcEngine(wb);
            var astNode = calcEngine.Parse(formula);
            Assert.Throws(Is.TypeOf<NotImplementedException>().With.Message.EqualTo(notSupportedMessage), () => astNode.Evaluate());
        }
    }
}
