using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Globalization;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class TextTests
    {
        [TestCase(@"ABCDEF123", @"ABCDEF123")]
        [TestCase(@"ァィゥェォッャュョヮ", @"ｧｨｩｪｫｯｬｭｮヮ")] // Small katakana, there is no half wa variant
        [TestCase(@"アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン", @"ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ")]
        [TestCase("！＂＃\uff04％＆＇（）＊\uff0b，－．／０１２３４５６７８９：；\uff1c\uff1d\uff1e？＠", @"!""#$%&'()*+,-./0123456789:;<=>?@")]
        [TestCase(@"ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ", @"ABCDEFGHIJKLMNOPQRSTUVWXYZ")]
        [TestCase("［＼］\uff3e＿\uff40ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛\uff5c｝\uff5e", @"[\]^_`abcdefghijklmnopqrstuvwxyz{|}~")]
        [TestCase(@"―‘’”、。「」゛゜・ー￥", @"ｰ`'""､｡｢｣ﾞﾟ･ｰ\")]
        public void Asc_converts_fullwidth_characters_to_halfwidth_characters(string input, string expected)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"ASC(\"{input}\")"), Is.EqualTo(expected));
        }

        [Test]
        public void Char_returns_error_on_empty_string()
        {
            // Calc engine tries to coerce it to number and fails. It never even reaches the functions.
            Assert.That(XLWorkbook.EvaluateExpr(@"CHAR("""")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase(0)]
        [TestCase(256)]
        [TestCase(9797)]
        public void Char_number_must_be_between_1_and_255(int number)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"CHAR({number})"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase(48, '0')]
        [TestCase(97, 'a')]
        [TestCase(128, '€')]
        [TestCase(138, 'Š')]
        [TestCase(169, '©')]
        [TestCase(182, '¶')]
        [TestCase(230, 'æ')]
        [TestCase(255, 'ÿ')]
        [TestCase(255.9, 'ÿ')]
        public void Char_interprets_number_as_win1252(double number, char expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"CHAR({number})");
            Assert.That(actual, Is.EqualTo(expected.ToString()));
        }

        [Test]
        public void Clean_empty_string_is_empty_string()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"CLEAN("""")"), Is.EqualTo(""));
        }

        [Test]
        public void Clean_removes_control_characters()
        {
            var actual = XLWorkbook.EvaluateExpr(@"CLEAN(CHAR(9)&""Monthly report""&CHAR(10))");
            Assert.That(actual, Is.EqualTo("Monthly report"));

            actual = XLWorkbook.EvaluateExpr(@"CLEAN(""   "")");
            Assert.That(actual, Is.EqualTo("   "));
        }

        [Test]
        public void Code_returns_error_on_empty_string()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"CODE("""")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("A", 65)]
        [TestCase("BCD", 66)]
        [TestCase("€", 128)]
        [TestCase("ÿ", 255)]
        public void Code_returns_win1252_codepoint_of_first_character(string text, int expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void Code_is_inverse_to_char()
        {
            for (var i = 1; i < 256; ++i)
                Assert.That(XLWorkbook.EvaluateExpr($"CODE(CHAR({i}))"), Is.EqualTo(i));
        }

        [TestCase("π")]
        [TestCase("ب")]
        [TestCase("😃")]
        [TestCase("♫")]
        [TestCase("ひ")]
        public void Code_returns_question_mark_code_on_non_win1252_chars(string text)
        {
            var expected = XLWorkbook.EvaluateExpr("CODE(\"?\")");
            var actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
            Assert.Multiple(() =>
            {
                Assert.That(expected, Is.EqualTo(63));
                Assert.That(actual, Is.EqualTo(expected));
            });
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void Concat_concatenates_scalar_values()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate(@"CONCAT(""ABC"",123,TRUE,IF(TRUE,),1.25)");
            Assert.That(actual, Is.EqualTo("ABC123TRUE1,25"));

            actual = ws.Evaluate(@"CONCAT("""",""123"")");
            Assert.That(actual, Is.EqualTo("123"));

            ws.FirstCell().SetValue(20.5)
                .CellBelow().SetValue("AB")
                .CellBelow().SetFormulaA1("DATE(2019,1,1)")
                .CellBelow().SetFormulaA1("CONCAT(A1:A3)");

            actual = ws.Cell("A4").Value;
            Assert.That(actual, Is.EqualTo("20,5AB43466"));
        }

        [Test]
        public void Concat_concatenates_array_values()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT({""A"",""B"",""C""},{0,1},{2;3},{4,5,6;7,8,9},""Z"")"), Is.EqualTo("ABC0123456789Z"));
        }

        [Test]
        public void Concat_concatenates_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("C2").InsertData(new object[]
            {
                ("A", "B", "C"),
                (1, 2, 3, 4),
                (5, 6, 7, 8),
            });
            Assert.That(ws.Evaluate("CONCAT(C2:E2,C3:F4,C2,\"Z\")"), Is.EqualTo("ABC12345678AZ"));
        }

        [Test]
        public void Concat_has_limit_of_32767_characters()
        {
            Assert.That(XLWorkbook.EvaluateExpr("CONCAT(REPT(\"A\",32768))"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Concat_accepts_only_area_references()
        {
            // Only areas are accepted, not unions
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate("CONCAT((C2:E2,C3:F4),C2,\"Z\")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Concat_propagates_error_values()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",#DIV/0!,5)"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",{""D"",#DIV/0!,7},5)"), Is.EqualTo(XLError.DivisionByZero));
            });

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B5").SetValue(XLError.DivisionByZero).CellBelow().SetValue(5);
            Assert.That(ws.Evaluate("CONCAT(\"ABC\",B5:B6)"), Is.EqualTo(XLError.DivisionByZero));
        }

        [Test]
        public void Concat_treats_blanks_as_empty_string()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",,""123"",)"), Is.EqualTo("ABC123"));
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void Concatenate_concatenates_scalar_values()
        {
            using var wb = new XLWorkbook();
            var actual = wb.Evaluate(@"CONCATENATE(""ABC"",123,4.56,IF(TRUE,),TRUE)");
            Assert.That(actual, Is.EqualTo("ABC1234,56TRUE"));

            actual = wb.Evaluate(@"CONCATENATE("""",""123"")");
            Assert.That(actual, Is.EqualTo("123"));
        }

        [Test]
        public void Concatenate_with_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = "Hello";
            ws.Cell("B1").Value = "World";
            ws.Cell("C1").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";
            ws.Cell("A3").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate(@"CONCATENATE(A1,"" "",B1)"), Is.EqualTo("Hello World"));

                // The result on C1 is on the same row (only one intersected cell) means implicit intersection
                // results in a one value per intersection and thus correct value. The A3 intersects two cells
                // and thus results in #VALUE! error.
                Assert.That(ws.Cell("C1").Value, Is.EqualTo("Hello World"));
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Concatenate_has_limit_of_32767_characters()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32767))"), Is.Not.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32768))"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Concatenate_uses_implicit_intersection_on_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue(20)
                .CellBelow().SetValue("AB")
                .CellBelow().SetFormulaA1("DATE(2019,1,1)");

            // Calling cell is 1st row, so formula should return A1
            ws.Cell("B1").SetFormulaA1("CONCATENATE(A1:A3)");
            Assert.That(ws.Cell("B1").Value, Is.EqualTo("20"));

            // Calling cell is 2nd row, so formula should return A2
            ws.Cell("B2").SetFormulaA1("CONCATENATE(A1:A3)");
            Assert.That(ws.Cell("B2").Value, Is.EqualTo("AB"));

            // Calling cell is 3rd row, so formula should return A3's textual representation
            ws.Cell("B3").SetFormulaA1("CONCATENATE(A1:A3)");
            Assert.That(ws.Cell("B3").Value, Is.EqualTo("43466"));

            // Calling cell doesn't share row with any cell in parameter range.
            ws.Cell("A4").SetFormulaA1("CONCATENATE(A1:A3)");
            Assert.That(ws.Cell("A4").Value, Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Dollar_coercion()
        {
            // Empty string is not coercible to number
            Assert.That(XLWorkbook.EvaluateExpr("DOLLAR(\"\", 3)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        // en-US culture differs between .NET Fx and Core for negative currency -> no test for negative
        [TestCase(123.54, 3, ExpectedResult = "$123.540")]
        [TestCase(123.54, 3.9, ExpectedResult = "$123.540")]
        [TestCase(1234.567, 2, ExpectedResult = "$1,234.57")]
        [TestCase(1250, -2, ExpectedResult = "$1,300")]
        [TestCase(1, -1E+100, ExpectedResult = "$0")]
        public string Dollar_en(double number, double decimals)
        {
            using var wb = new XLWorkbook();
            return wb.Evaluate($"DOLLAR({number},{decimals})").GetText();
        }

        [SetCulture("cs-CZ")]
        [TestCase(123.54, 3, ExpectedResult = "123,540 Kč")]
        [TestCase(-1234.567, 4, ExpectedResult = "-1 234,5670 Kč")]
        [TestCase(-1250, -2, ExpectedResult = "-1 300 Kč")]
        public string Dollar_cs(double number, double decimals)
        {
            using var wb = new XLWorkbook();
            var formula = $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
            return wb.Evaluate(formula).GetText();
        }

        [SetCulture("de-DE")]
        [TestCase(1234.567, 2, ExpectedResult = "1.234,57 €")]
        [TestCase(1234.567, -2, ExpectedResult = "1.200 €")]
        [TestCase(-1234.567, 4, ExpectedResult = "-1.234,5670 €")]
        public string Dollar_de(double number, double decimals)
        {
            using var wb = new XLWorkbook();
            var formula = $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
            return wb.Evaluate(formula).GetText();
        }

        [Test]
        public void Dollar_uses_two_decimal_places_by_default()
        {
            using var wb = new XLWorkbook();
            var actual = wb.Evaluate("DOLLAR(123.543)");
            Assert.That(actual, Is.EqualTo("$123.54"));
        }

        [Test]
        public void Dollar_can_have_at_most_127_decimal_places()
        {
            using var wb = new XLWorkbook();
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("DOLLAR(1,99)"), Is.EqualTo("$1." + new string('0', 99)));
                Assert.That(wb.Evaluate("DOLLAR(1,128)"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Exact_Empty_Input_String()
        {
            object actual = XLWorkbook.EvaluateExpr(@"Exact("""", """")");
            Assert.That(actual, Is.EqualTo(true));
        }

        [Test]
        public void Exact_Value()
        {
            object actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""asdf"")");
            Assert.That(actual, Is.EqualTo(true));

            actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""ASDF"")");
            Assert.That(actual, Is.EqualTo(false));

            actual = XLWorkbook.EvaluateExpr(@"Exact(123, 123)");
            Assert.That(actual, Is.EqualTo((true)));

            actual = XLWorkbook.EvaluateExpr(@"Exact(321, 123)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void Find_Empty_Pattern_And_Empty_Text()
        {
            Assert.Multiple(() =>
            {
                // Different behavior from SEARCH
                Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", """")"), Is.EqualTo(1));

                Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", ""a"", 2)"), Is.EqualTo(2));
            });
        }

        [Test]
        public void Find_Empty_Search_Pattern_Returns_Start_Of_Text()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", ""asdf"")"), Is.EqualTo(1));
        }

        [Test]
        public void Find_Looks_Only_From_Start_Position_Onward()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""This"", ""This is some text"", 2)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Start_Position_Too_Large()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""abc"", ""abcdef"", 10)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Start_Position_Too_Small()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""text"", ""This is some text"", 0)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Empty_Searched_Text_Returns_Error()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""abc"", """")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_String_Not_Found()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""123"", ""asdf"")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Case_Sensitive_String_Not_Found()
        {
            // Find is case-sensitive
            Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""excel"", ""Microsoft Excel 2010"")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Value()
        {
            var actual = XLWorkbook.EvaluateExpr(@"FIND(""Tuesday"", ""Today is Tuesday"")");
            Assert.That(actual, Is.EqualTo(10));

            // Doesnt support wildcards
            actual = XLWorkbook.EvaluateExpr(@"FIND(""T*y"", ""Today is Tuesday"")");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Find_Arguments_Are_Converted_To_Expected_Types()
        {
            var actual = XLWorkbook.EvaluateExpr(@"FIND(1.2, ""A1.2B"")");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr(@"FIND(TRUE, ""ATRUE"")");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr(@"FIND(23, 1.2345)");
            Assert.That(actual, Is.EqualTo(3));

            actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""aaaaa"", ""2 1/2"")");
            Assert.That(actual, Is.EqualTo(2));
        }

        [Test]
        public void Find_Error_Arguments_Return_The_Error()
        {
            var actual = XLWorkbook.EvaluateExpr(@"FIND(#N/A, ""a"")");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));

            actual = XLWorkbook.EvaluateExpr(@"FIND("""", #N/A)");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));

            actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""a"", #N/A)");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));
        }

        [Test]
        public void Fixed_coercion()
        {
            using var wb = new XLWorkbook();
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("""FIXED("asdf")"""), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(wb.Evaluate("""FIXED(1234,1,"TRUE")"""), Is.EqualTo("1234.0"));
                Assert.That(wb.Evaluate("""FIXED(1234,1,"FALSE")"""), Is.EqualTo("1,234.0"));
                Assert.That(wb.Evaluate("""FIXED(1234,1,"0")"""), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Fixed_examples()
        {
            using var wb = new XLWorkbook();
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("FIXED(1234567)"), Is.EqualTo("1,234,567.00"));
                Assert.That(wb.Evaluate("FIXED(1234567.555555,4,TRUE)"), Is.EqualTo("1234567.5556"));
                Assert.That(wb.Evaluate("FIXED(.555555,10)"), Is.EqualTo("0.5555550000"));
                Assert.That(wb.Evaluate("FIXED(1234567,-3)"), Is.EqualTo("1,235,000"));
            });
        }

        [Test]
        public void Fixed_en()
        {
            var actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,4)");
            Assert.That(actual, Is.EqualTo("17,300.6700"));

            actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,2,TRUE)");
            Assert.That(actual, Is.EqualTo("17300.67"));

            actual = XLWorkbook.EvaluateExpr("FIXED(17300.67)");
            Assert.That(actual, Is.EqualTo("17,300.67"));

            actual = XLWorkbook.EvaluateExpr("FIXED(1,-1E+300)");
            Assert.That(actual, Is.EqualTo("0"));
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void Fixed_cs()
        {
            using var wb = new XLWorkbook();
            var actual = wb.Evaluate("FIXED(17300.67,4)");
            Assert.That(actual, Is.EqualTo("17 300,6700"));

            actual = wb.Evaluate("FIXED(17300.67,2,TRUE)");
            Assert.That(actual, Is.EqualTo("17300,67"));

            actual = wb.Evaluate("FIXED(17300.67)");
            Assert.That(actual, Is.EqualTo("17 300,67"));
        }

        [Test]
        public void Fixed_can_have_at_most_127_decimal_places()
        {
            using var wb = new XLWorkbook();
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("FIXED(1,99)"), Is.EqualTo("1." + new string('0', 99)));
                Assert.That(wb.Evaluate("FIXED(1,128)"), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Left_returns_whole_text_when_requested_length_is_greater_than_text_length()
        {
            var actual = XLWorkbook.EvaluateExpr(@"LEFT(""ABC"", 5)");
            Assert.That(actual, Is.EqualTo("ABC"));
        }

        [Test]
        public void Left_takes_one_character_by_default()
        {
            var actual = XLWorkbook.EvaluateExpr("""LEFT("ABC")""");
            Assert.That(actual, Is.EqualTo("A"));
        }

        [Test]
        public void Left_returns_error_on_negative_number_of_chars()
        {
            Assert.That(XLWorkbook.EvaluateExpr("""LEFT("ABC", -1)"""), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Left_returns_empty_string_on_empty_input()
        {
            var actual = XLWorkbook.EvaluateExpr("""LEFT("")""");
            Assert.That(actual, Is.EqualTo(""));
        }

        [TestCase("ABC", 2, ExpectedResult = "AB")]
        [TestCase("ABC", 2.9, ExpectedResult = "AB")]
        [TestCase("ABC", 3, ExpectedResult = "ABC")]
        [TestCase("\uD83D\uDC69Z", 1, ExpectedResult = "\uD83D\uDC69")] // Paired surrogate
        [TestCase("\uD83D\uDC69Z", 2, ExpectedResult = "\uD83D\uDC69Z")] // Paired surrogate
        public string Left_takes_specified_number_of_characters(string text, double numChars)
        {
            return XLWorkbook.EvaluateExpr($"""LEFT("{text}", {numChars})""").GetText();
        }

        [TestCase("", ExpectedResult = 0)]
        [TestCase("word", ExpectedResult = 4)]
        [TestCase("A\r\n", ExpectedResult = 3)]
        [TestCase("H", ExpectedResult = 1)]
        [TestCase("\ud83d\ude0a", ExpectedResult = 2)] // Smile emoji
        [TestCase("Smile: \ud83d\ude0a!", ExpectedResult = 10)] // Smile emoji
        public double Len_returns_number_of_code_units(string text)
        {
            return XLWorkbook.EvaluateExpr($"""LEN("{text}")""").GetNumber();
        }

        [SetCulture("en-US")]
        [TestCase("", ExpectedResult = "")]
        [TestCase("ABC", ExpectedResult = "abc")]
        [TestCase("Intelligence 2.0!", ExpectedResult = "intelligence 2.0!")]
        [TestCase("ͶꝎＫǢ", ExpectedResult = "ͷꝏｋǣ")] // Converts even non-latin chars
        [TestCase("Σ SUM Σ end Σ", ExpectedResult = "σ sum σ end ς")] // Bug for bug behavior of Excel. Σ at the end is turned to ς
        public string Lower_en(string text)
        {
            using var wb = new XLWorkbook();
            return wb.Evaluate($"""LOWER("{text}")""").GetText();
        }

        [SetCulture("tr-TR")]
        [TestCase("INTELLIGENCE 2.0!", ExpectedResult = "ıntellıgence 2.0!")] // Turkey converts I to i without dot
        [TestCase("ΣΣΣΣ", ExpectedResult = "σσσς")]
        public string Lower_tr(string text)
        {
            using var wb = new XLWorkbook();
            return wb.Evaluate($"""LOWER("{text}")""").GetText();
        }

        [Test]
        public void Mid_returns_rest_of_text_when_end_is_out_of_text_bounds()
        {
            var actual = XLWorkbook.EvaluateExpr("""MID("ABC",1,5)""");
            Assert.That(actual, Is.EqualTo("ABC"));
        }

        [Test]
        public void Mid_when_start_is_after_end_of_text_return_empty_string()
        {
            var actual = XLWorkbook.EvaluateExpr("""MID("ABC",5,5)""");
            Assert.That(actual, Is.EqualTo(""));
        }

        [TestCase(0.9)]
        [TestCase(0)]
        [TestCase(-5)]
        [TestCase(int.MaxValue + 1d)]
        [TestCase(int.MaxValue + 5d)]
        public void Mid_start_must_be_at_least_one_and_at_most_max_int(double start)
        {
            var actual = XLWorkbook.EvaluateExpr($"""MID("ABC",{start},1)""");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase(-0.1)]
        [TestCase(-5)]
        [TestCase(int.MaxValue + 1d)]
        [TestCase(int.MaxValue + 5d)]
        public void Mid_length_must_be_at_least_zero_and_at_most_max_int(double length)
        {
            var actual = XLWorkbook.EvaluateExpr($"""MID("ABC",1,{length})""");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("", 1, 1, ExpectedResult = "")]
        [TestCase("ABC", 2, 2, ExpectedResult = "BC")]
        [TestCase("ABC", 2, 0, ExpectedResult = "")]
        [TestCase("ABC", 3, 5, ExpectedResult = "")]
        [TestCase(@"abcdef", 3, 2, ExpectedResult = "cd")]
        [TestCase(@"abcdef", 4, 5, ExpectedResult = "def")]
        public string Mid_returns_substring(string text, double start, double length)
        {
            return XLWorkbook.EvaluateExpr($"""MID("{text}",{start},{length})""").GetText();
        }

        [Test]
        public void Mid_uses_code_units()
        {
            Assert.Multiple(() =>
            {
                // MID returns unpaired surrogates
                Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,3)"""), Is.EqualTo("😊\uD83D"));
                Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,4)"""), Is.EqualTo("😊😊"));
                Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",2,4)"""), Is.EqualTo("\uDE0A😊\uD83D"));
                Assert.That(XLWorkbook.EvaluateExpr("""LEN(MID("😊😊😊",1,3))"""), Is.EqualTo(3));
            });
        }

        [TestCase("", 0d)]
        [TestCase("+ 1", 1d)]
        [TestCase("+1", 1d)]
        [TestCase("+1.23", 1.23)]
        [TestCase("- 1.23", -1.23)]
        [TestCase(" - 0 1 2 . 3 4 ", -12.34)]
        [TestCase(" - 0 \t1\t2\r .\n3 4 ", -12.34)]
        [TestCase(".1", 0.1)]
        [TestCase("-.1", -0.1)]
        [TestCase("1.234567890E+307", 1.234567890E+307)]
        [TestCase("1.234567890E-307", 1.234567890E-307d)]
        [TestCase("1.234567890E-309", 0d)]
        [TestCase("-1.234567890E-307", -1.234567890E-307d)]
        [TestCase(".99999999999999", 0.99999999999999)]
        [TestCase("1,23,4", 1234)]
        [TestCase("1,234,56", 123456)]
        [TestCase("1e-308", 0)]
        [TestCase("-1e-308", 0)]
        [TestCase("75825%", 758.25)]
        [TestCase("75825%%", 7.5825)]
        [TestCase("(56.4)", -56.4)]
        [TestCase("(128)%", -1.28)]
        public void NumberValue_converts_text_to_number(string text, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExprCurrent($"NUMBERVALUE(\"{text}\")");
            Assert.That(actual, Is.EqualTo(expectedResult));
        }

        [Test]
        [SetCulture("de-DE")]
        public void NumberValue_takes_separators_from_current_culture()
        {
            var actual = (double)XLWorkbook.EvaluateExprCurrent("NUMBERVALUE(\"10.0.00.0,25\")");
            Assert.That(actual, Is.EqualTo(100000.25));
        }

        [TestCase("1,234.56", ".", ",", 1234.56d)]
        [TestCase("1.234,56", ",", ".", 1234.56d)]
        [TestCase("1.234,56", ",ABC", ".DEF", 1234.56d)] // Only first char of separators is used
        public void NumberValue_optional_parameters_can_set_decimal_and_group_separators(string text, string @decimal, string group, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"NUMBERVALUE(\"{text}\",\"{@decimal}\",\"{group}\")");
            Assert.That(actual, Is.EqualTo(expectedResult));
        }

        [TestCase("NUMBERVALUE(\"123.45\", \".\", \".\")")] // Group separator same as decimal separator
        [TestCase("NUMBERVALUE(\"1.234.5\")")] // Two decimal separators
        [TestCase("NUMBERVALUE(\"1.234,5\")")] // Decimal separator before group separator
        [TestCase("NUMBERVALUE(\"12;34\")")] // Illegal character
        [TestCase("NUMBERVALUE(\"--1\")")] // Two minuses
        [TestCase("NUMBERVALUE(\"1.234567890E+308\")")] // Too large
        [TestCase("NUMBERVALUE(\"-1.234567890E+308\")")] // Too large (negative)
        [TestCase("NUMBERVALUE(\"1.234567890E-310\")")] // Too tiny
        [TestCase("NUMBERVALUE(\"-1.234567890E-310\")")] // Too tiny (negative)
        [TestCase("NUMBERVALUE(\"1\",\".\",\"\")")] // Empty group separator
        [TestCase("NUMBERVALUE(\"1\",\"\",\",\")")] // Empty decimal separators
        public void NumberValue_returns_error_on_unparsable_texts_out_of_range(string expression)
        {
            Assert.That(XLWorkbook.EvaluateExpr(expression), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("", ExpectedResult = "")]
        [TestCase("12aBC d123aD#$%sd^", ExpectedResult = "12Abc D123Ad#$%Sd^")]
        [TestCase("this is a TITLE", ExpectedResult = "This Is A Title")]
        [TestCase("2-way street", ExpectedResult = "2-Way Street")]
        [TestCase("76BudGet", ExpectedResult = "76Budget")]
        [TestCase("my name is francois botha", ExpectedResult = "My Name Is Francois Botha")]
        [TestCase("\ud83a\udd32", ExpectedResult = "\ud83a\udd32")] // U+1E932 has uppercase variant, but nothing changes, because PROPER uses code units
        public string Proper_upper_cases_first_letter_and_lower_cases_next_letters(string text)
        {
            return XLWorkbook.EvaluateExpr($"""PROPER("{text}")""").GetText();
        }

        [TestCase(1, 1)]
        [TestCase(1, 0)]
        [TestCase(1, 10)]
        [TestCase(10, 1)]
        [TestCase(10, 10)]
        public void Replace_beyond_limit_appends_replacement(int startPos, int length)
        {
            var actual = XLWorkbook.EvaluateExpr($"""REPLACE("",{startPos},{length},"new text")""");
            Assert.That(actual, Is.EqualTo("new text"));
        }

        [TestCase("Here is some obsolete text to replace.", 14, 13, "new text", ExpectedResult = "Here is some new text to replace.")]
        [TestCase("ABC", 1, 2, "D", ExpectedResult = "DC")]
        [TestCase("ABC", 3, 1, "D", ExpectedResult = "ABD")]
        [TestCase("ABC", 3, 0, "D", ExpectedResult = @"ABDC")]
        [TestCase("ABC", 4, 1, "D", ExpectedResult = @"ABCD")]
        [TestCase("ABC", 4, 0, "D", ExpectedResult = @"ABCD")]
        [TestCase("ABC", 1, 3, "D", ExpectedResult = "D")]
        [TestCase("ABC", 2, 2, "D", ExpectedResult = "AD")]
        [TestCase("ABC", 2, 0, "D", ExpectedResult = @"ADBC")]
        [TestCase("ABC", 2, 3, "D", ExpectedResult = "AD")]
        [TestCase(@"abcdefghijk", 3, 4, "XY", ExpectedResult = @"abXYghijk")]
        [TestCase(@"abcdefghijk", 3, 1, "12345", ExpectedResult = @"ab12345defghijk")]
        [TestCase(@"abcdefghijk", 15, 4, "XY", ExpectedResult = @"abcdefghijkXY")]
        public string Replace_replaces_value(string text, double startPos, int length, string replacement)
        {
            return XLWorkbook.EvaluateExpr($"""REPLACE("{text}",{startPos},{length},"{replacement}")""").GetText();
        }

        [Test]
        public void Replace_start_position_must_be_from_1_to_32767()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"D")"""), Is.EqualTo(@"DABC"));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",0.9,0,"D")"""), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",-1,0,"D")"""), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32767.9,"D")"""), Is.EqualTo("D"));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32768,"D")"""), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Replace_length_must_be_from_0_to_32767()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"")"""), Is.EqualTo("ABC"));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,-0.1,"D")"""), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32767.9,"D")"""), Is.EqualTo("D"));
                Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32768,"D")"""), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [Test]
        public void Rept_returns_empty_string_when_text_is_empty_string()
        {
            var actual = XLWorkbook.EvaluateExpr("""REPT("",3)""");
            Assert.That(actual, Is.EqualTo(""));
        }

        [TestCase(-1)]
        [TestCase(-0.1)]
        [TestCase(2147483648)]
        public void Rept_returns_error_when_count_is_negative_or_greater_than_max_int(double count)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"""REPT("",{count})"""), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Rept_limits_output_text_length_to_32767()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("""REPT("A",32767)"""), Is.EqualTo(new string('A', 32767)));
                Assert.That(XLWorkbook.EvaluateExpr("""REPT("A",32768)"""), Is.EqualTo(XLError.IncompatibleValue));
            });
        }

        [TestCase("ABC", 3, ExpectedResult = @"ABCABCABC")]
        [TestCase("123", 2.5, ExpectedResult = "123123")]
        [TestCase("Francois", 0, ExpectedResult = "")]
        [TestCase("Francois Botha,", 3, ExpectedResult = "Francois Botha,Francois Botha,Francois Botha,")]
        public string Rept_Value(string text, double count)
        {
            return XLWorkbook.EvaluateExpr($"""REPT("{text}",{count})""").GetText();
        }

        [TestCase(5)]
        [TestCase(3)]
        public void Right_returns_whole_text_when_requested_length_is_greater_than_text_length(int length)
        {
            var actual = XLWorkbook.EvaluateExpr($"""RIGHT("ABC",{length})""");
            Assert.That(actual, Is.EqualTo("ABC"));
        }

        [Test]
        public void Right_takes_one_character_by_default()
        {
            var actual = XLWorkbook.EvaluateExpr("""RIGHT("ABC")""");
            Assert.That(actual, Is.EqualTo("C"));
        }

        [Test]
        public void Right_returns_error_on_negative_number_of_chars()
        {
            Assert.That(XLWorkbook.EvaluateExpr("""RIGHT("ABC",-1)"""), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Right_returns_empty_string_on_empty_input()
        {
            var actual = XLWorkbook.EvaluateExpr("""RIGHT("")""");
            Assert.That(actual, Is.EqualTo(""));
        }

        [TestCase("ABC", 0, ExpectedResult = "")]
        [TestCase("ABC", 1, ExpectedResult = "C")]
        [TestCase("ABC", 2, ExpectedResult = "BC")]
        [TestCase("ABC", 3, ExpectedResult = "ABC")]
        [TestCase("ABC", 4, ExpectedResult = "ABC")]
        [TestCase("ABC", 2.9, ExpectedResult = "BC")]
        [TestCase("Z\uD83D\uDC69", 1, ExpectedResult = "\uD83D\uDC69")] // Smiley emoji
        [TestCase("\uD83D\uDC69Z", 2, ExpectedResult = "\uD83D\uDC69Z")]
        [TestCase("\uD83D\uDC69Z", 3, ExpectedResult = "\uD83D\uDC69Z")]
        public string Right_takes_specified_number_of_characters(string text, double numChars)
        {
            return XLWorkbook.EvaluateExpr($"""RIGHT("{text}",{numChars})""").GetText();
        }

        [Test]
        public void Search_Empty_Pattern_And_Empty_Text()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH("""", """")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Empty_Search_Pattern_Returns_Start_Of_Text()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", ""asdf"")");
            Assert.That(actual, Is.EqualTo(1));
        }

        [Test]
        public void Search_Looks_Only_From_Start_Position_Onward()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""This"", ""This is some text"", 2)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Start_Position_Too_Large()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", ""abcdef"", 10)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Start_Position_Too_Small()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 0)"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Empty_Searched_Text_Returns_Error()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", """")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Text_Not_Found()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""123"", ""asdf"")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Wildcard_String_Not_Found()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""soft?2010"", ""Microsoft Excel 2010"")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        // http://www.excel-easy.com/examples/find-vs-search.html
        [Test]
        public void Search_Value()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Tuesday"", ""Today is Tuesday"")");
            Assert.That(actual, Is.EqualTo(10));

            // The search is case-insensitive
            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""excel"", ""Microsoft Excel 2010"")");
            Assert.That(actual, Is.EqualTo(11));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""soft*2010"", ""Microsoft Excel 2010"")");
            Assert.That(actual, Is.EqualTo(6));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Excel 20??"", ""Microsoft Excel 2010"")");
            Assert.That(actual, Is.EqualTo(11));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 14)");
            Assert.That(actual, Is.EqualTo(14));
        }

        [Test]
        public void Search_Tilde_Escapes_Next_Char()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SEARCH(""~a~b~"", ""ab"")");
            Assert.That(actual, Is.EqualTo(1));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""a*"")");
            Assert.That(actual, Is.EqualTo(1));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""ab"")");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""a?"")");
            Assert.That(actual, Is.EqualTo(1));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""ab"")");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Search_Arguments_Are_Converted_To_Expected_Types()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SEARCH(1.2, ""A1.2B"")");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(TRUE, ""ATRUE"")");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(23, 1.2345)");
            Assert.That(actual, Is.EqualTo(3));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""aaaaa"", ""2 1/2"")");
            Assert.That(actual, Is.EqualTo(2));
        }

        [Test]
        public void Search_Error_Arguments_Return_The_Error()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SEARCH(#N/A, ""a"")");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", #N/A)");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));

            actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""a"", #N/A)");
            Assert.That(actual, Is.EqualTo(XLError.NoValueAvailable));
        }

        [Test]
        public void Substitute_replaces_n_th_occurence()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday."", ""Tuesday"", ""Monday"")");
            Assert.That(actual, Is.EqualTo("This is a Monday."));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 1)");
            Assert.That(actual, Is.EqualTo("This is a Monday. Next week also has a Tuesday."));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 2)");
            Assert.That(actual, Is.EqualTo("This is a Tuesday. Next week also has a Monday."));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", """", ""Monday"")");
            Assert.That(actual, Is.EqualTo("This is a Tuesday. Next week also has a Tuesday."));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", """")");
            Assert.That(actual, Is.EqualTo("This is a . Next week also has a ."));
        }

        [Test]
        public void Substitute_on_empty_string_returns_empty_string()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE("""","""",""Monday"")");
            Assert.That(actual, Is.EqualTo(""));
        }

        [Test]
        public void Substitute_is_case_sensitive()
        {
            var actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("A","a","Z")""");
            Assert.That(actual, Is.EqualTo("A"));
        }

        [Test]
        public void Substitute_returns_original_string_when_occurrence_is_not_found()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABCABC"",""A"",""Z"",3)");
            Assert.That(actual, Is.EqualTo(@"ABCABC"));
        }

        [Test]
        public void Substitute_searches_for_every_occurence()
        {
            // AA is matches at every character, it doesn't skip
            var actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("AAAAAAAA","AA","ZZ",3)""");
            Assert.That(actual, Is.EqualTo(@"AAZZAAAA"));
        }

        [Test]
        public void Substitute_occurence_must_be_between_one_and_max_int()
        {
            var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"",0.9)");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483646.9)");
            Assert.That(actual, Is.EqualTo("ABC"));

            actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483647)");
            Assert.That(actual, Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void T_returns_empty_string_on_non_text()
        {
            var actual = XLWorkbook.EvaluateExpr("T(TODAY())");
            Assert.That(actual, Is.EqualTo(""));

            actual = XLWorkbook.EvaluateExpr("T(IF(TRUE,,))");
            Assert.That(actual, Is.EqualTo(""));

            actual = XLWorkbook.EvaluateExpr("T(TRUE)");
            Assert.That(actual, Is.EqualTo(""));

            actual = XLWorkbook.EvaluateExpr("T(123)");
            Assert.That(actual, Is.EqualTo(""));
        }

        [Test]
        public void T_propagates_error()
        {
            Assert.That(XLWorkbook.EvaluateExpr("T(#DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
        }

        [Test]
        public void T_returns_text_when_value_is_text()
        {
            var actual = XLWorkbook.EvaluateExpr("""T("asdf")""");
            Assert.That(actual, Is.EqualTo("asdf"));

            actual = XLWorkbook.EvaluateExpr("""T("")""");
            Assert.That(actual, Is.EqualTo(""));
        }

        [Test]
        public void T_returns_array_of_results_when_argument_is_array()
        {
            const string formula = """T({"A",5,"B"})""";
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr($"""COLUMNS({formula})"""), Is.EqualTo(3));
                Assert.That(XLWorkbook.EvaluateExpr($"""ROWS({formula})"""), Is.EqualTo(1));
                Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,1)"""), Is.EqualTo("A"));
                Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,2)"""), Is.EqualTo(""));
                Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,3)"""), Is.EqualTo("B"));

                // Array doesn't propagate single error, but returns errors in the array
                Assert.That(XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,1)"""), Is.EqualTo("A"));
                Assert.That(XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,2)"""), Is.EqualTo(XLError.CellReference));
            });
        }

        [Test]
        public void T_returns_text_of_first_cell_in_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = "ABC";
            ws.Cell("B4").Value = 10;
            ws.Cell("B5").Value = XLError.NoValueAvailable;

            Assert.Multiple(() =>
            {
                Assert.That(ws.Evaluate("T(B3:B4)"), Is.EqualTo("ABC"));
                Assert.That(ws.Evaluate("TYPE(T(B3:B4))"), Is.EqualTo(2)); // Is text, not array

                Assert.That(ws.Evaluate("T(B4:C4)"), Is.EqualTo(""));

                Assert.That(ws.Evaluate("T(B5:C5)"), Is.EqualTo(XLError.NoValueAvailable));
            });
        }

        [Test]
        public void Text_returns_empty_string_on_empty_string()
        {
            var actual = XLWorkbook.EvaluateExpr(@"TEXT(1913415.93,"""")");
            Assert.That(actual, Is.EqualTo(""));
        }

        [TestCase("DATE(2010, 1, 1)", "yyyy-MM-dd", ExpectedResult = "2010-01-01")]
        [TestCase("1469.07", "0,000,000.00", ExpectedResult = "0,001,469.07")]
        [TestCase("1913415.93", "#,000.00", ExpectedResult = "1,913,415.93")]
        [TestCase("2800", "$0.00", ExpectedResult = "$2800.00")]
        [TestCase("0.4", "0%", ExpectedResult = "40%")]
        [TestCase("DATE(2010, 1, 1)", "MMMM yyyy", ExpectedResult = "January 2010")]
        [TestCase("DATE(2010, 1, 1)", "M/d/y", ExpectedResult = "1/1/10")]
        [TestCase("1234.567", "$0.00", ExpectedResult = "$1234.57")]
        [TestCase(".125", "$0.0%", ExpectedResult = "$12.5%")]
        [TestCase("1234.567", "YYYY-MM-DD HH:MM:SS", ExpectedResult = "1903-05-18 13:36:28")] // Excel is one second off (29), but that is in the library
        [TestCase("\"0.0245\"", "00%", ExpectedResult = "02%")]
        public string Text_formats_number(string numberArg, string format)
        {
            return XLWorkbook.EvaluateExpr($"TEXT({numberArg},\"{format}\")").GetText();
        }

        [TestCase("\"211x\"", ExpectedResult = "211x")]
        [TestCase("true", ExpectedResult = "TRUE")]
        public string Text_returns_string_representation_of_non_numbers(string valueArg)
        {
            return XLWorkbook.EvaluateExpr($@"TEXT({valueArg},""#00"")").GetText();
        }

        [TestCase(2020, 11, 1, 9, 23, 11, "m/d/yyyy h:mm:ss", "11/1/2020 9:23:11")]
        [TestCase(2023, 7, 14, 2, 12, 3, "m/d/yyyy h:mm:ss", "7/14/2023 2:12:03")]
        [TestCase(2025, 10, 14, 2, 48, 55, "m/d/yyyy h:mm:ss", "10/14/2025 2:48:55")]
        [TestCase(2023, 2, 19, 22, 1, 38, "m/d/yyyy h:mm:ss", "2/19/2023 22:01:38")]
        [TestCase(2025, 12, 19, 19, 43, 58, "m/d/yyyy h:mm:ss", "12/19/2025 19:43:58")]
        [TestCase(2034, 11, 16, 1, 48, 9, "m/d/yyyy h:mm:ss", "11/16/2034 1:48:09")]
        [TestCase(2018, 12, 10, 11, 22, 42, "m/d/yyyy h:mm:ss", "12/10/2018 11:22:42")]
        public void Text_formats_serial_dates(int year, int months, int days, int hour, int minutes, int seconds, string format, string expected)
        {
            Assert.That(XLWorkbook.EvaluateExpr($@"TEXT(DATE({year},{months},{days}) + TIME({hour},{minutes},{seconds}),""{format}"")"), Is.EqualTo(expected));
        }

        [Test]
        public void Text_propagates_errors()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"TEXT(#REF!,""#00"")"), Is.EqualTo(XLError.CellReference));
        }

        [TestCase("TEXTJOIN(\",\",TRUE,A1:B2)", "A,B,D")]
        [TestCase("TEXTJOIN(\",\",FALSE,A1:B2)", "A,,B,D")]
        [TestCase("TEXTJOIN(\",\",FALSE,A1,A2,B1,B2)", "A,B,,D")]
        [TestCase("TEXTJOIN(\",\",FALSE,1)", "1")]
        [TestCase("TEXTJOIN(\",\", TRUE, A:A, B:B)", "A,B,D")]
        [TestCase("TEXTJOIN(\",\", TRUE, D1:E2)", "")]
        [TestCase("TEXTJOIN(\",\", FALSE, D1:E2)", ",,,")]
        [TestCase("TEXTJOIN(\",\", FALSE, D1:D32768)", ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,")]
        [TestCase("TEXTJOIN(0, FALSE, A1:B2)", "A00B0D")]
        [TestCase("TEXTJOIN(false, FALSE, A1:B2)", @"AFALSEFALSEBFALSED")]
        [TestCase("TEXTJOIN(\",\", 0, A1:B2)", "A,,B,D")]
        [TestCase("TEXTJOIN(\",\", 100, A1:B2)", "A,B,D")]
        [TestCase("TEXTJOIN(B2, FALSE, A1:B2)", @"ADDBDD")]
        [TestCase("TEXTJOIN(\",\", FALSE, 12345.67, DATE(2018, 10, 30))", "12345.67,43403")]
        [TestCase("TEXTJOIN(\",\", \"FALSE\", A1:B2)", "A,,B,D")]
        public void TextJoin_joins_arguments_with_specified_delimiter(string formula, string expectedOutput)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "A";
            ws.Cell("A2").Value = "B";
            ws.Cell("B1").Value = "";
            ws.Cell("B2").Value = "D";

            ws.Cell("C1").FormulaA1 = formula;
            var a = ws.Cell("C1").Value;

            Assert.That(a, Is.EqualTo(expectedOutput));
        }

        [TestCase("TEXTJOIN(\",\", FALSE, D1:D32769)")]
        public void TextJoin_output_can_be_at_most_32767(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("C1").FormulaA1 = formula;

            // Excel actually returns #CALC!, but we don't have that error, mostly
            // because parser doesn't recognize it.
            Assert.That(ws.Cell("C1").Value, Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("TEXTJOIN(\",\", \"Invalid\", \"Hello\", \"World\")")]
        public void TextJoin_coercion(string formula)
        {
            Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("", ExpectedResult = "")]
        [TestCase(" ", ExpectedResult = "")]
        [TestCase("    ", ExpectedResult = "")]
        [TestCase(" Break\r\n   Line   ", ExpectedResult = "Break\r\n Line")]
        [TestCase("non-whitespace-text", ExpectedResult = "non-whitespace-text")]
        [TestCase("white space text", ExpectedResult = "white space text")]
        [TestCase(" some text with padding   ", ExpectedResult = "some text with padding")]
        [TestCase(" \t  A  \t ", ExpectedResult = "\t A \t")]
        public string Trim_trims_spaces_and_removes_multi_spaces_from_inside_text(string text)
        {
            return XLWorkbook.EvaluateExpr($"""TRIM("{text}")""").GetText();
        }

        [Test]
        public void Upper_empty_string_returns_empty_string()
        {
            Assert.That(XLWorkbook.EvaluateExpr("""UPPER("")"""), Is.EqualTo(""));
        }

        [Test]
        public void Upper_converts_text_to_upper_case()
        {
            var actual = XLWorkbook.EvaluateExpr("""UPPER("AbCdEfG")""");
            Assert.That(actual, Is.EqualTo(@"ABCDEFG"));
        }

        [SetCulture("tr-TR")]
        [Test]
        public void Upper_uses_workbook_culture()
        {
            // Türkiye converts i to İ, not I.
            using var wb = new XLWorkbook();
            Assert.That(wb.Evaluate("""UPPER("intelligence 2.0!")"""), Is.EqualTo("İNTELLİGENCE 2.0!"));
        }

        [Test]
        public void Value_Input_String_Is_Not_A_Number()
        {
            Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(""asdf"")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Value_FromBlankIsZero()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.Evaluate("VALUE(A1)"), Is.EqualTo(0d));
        }

        [Test]
        public void Value_FromEmptyStringIsError()
        {
            Assert.That(XLWorkbook.EvaluateExpr("VALUE(\"\")"), Is.EqualTo(XLError.IncompatibleValue));
        }

        [Test]
        public void Value_PassingUnexpectedTypes()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(14)"), Is.EqualTo(14d));
                Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(TRUE)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(FALSE)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(#DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
            });
        }

        [Test]
        public void Value_Value()
        {
            using var wb = new XLWorkbook();

            Assert.Multiple(() =>
            {
                // Examples from spec
                Assert.That(wb.Evaluate("VALUE(\"123.456\")"), Is.EqualTo(123.456d));
                Assert.That(wb.Evaluate("VALUE(\"$1,000\")"), Is.EqualTo(1000d));
                Assert.That(wb.Evaluate("VALUE(\"23-Mar-2002\")"), Is.EqualTo(new DateTime(2002, 3, 23).ToSerialDateTime()));
                Assert.That((double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")"), Is.EqualTo(0.188056d).Within(0.000001d));
            });
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void Value_NonEnglish()
        {
            using var wb = new XLWorkbook();

            Assert.Multiple(() =>
            {
                // Examples from spec
                Assert.That(wb.Evaluate("VALUE(\"123,456\")"), Is.EqualTo(123.456d));
                Assert.That(wb.Evaluate("VALUE(\"1 000 Kč\")"), Is.EqualTo(1000d));
                Assert.That(wb.Evaluate("VALUE(\"23-bře-2002\")"), Is.EqualTo(37338d));
                Assert.That((double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")"), Is.EqualTo(0.188056d).Within(0.000001d));

                // Various number/currency formats
                Assert.That(wb.Evaluate("VALUE(\"(1)\")"), Is.EqualTo(-1d));
                Assert.That(wb.Evaluate("VALUE(\"(100%)\")"), Is.EqualTo(-1d));
            });
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("VALUE(\"(100%)\")"), Is.EqualTo(-1d));
                Assert.That(wb.Evaluate("VALUE(\"(1,5e1 Kč)\")"), Is.EqualTo(-15d));
                Assert.That(wb.Evaluate("VALUE(\"(1,5e3%)\")"), Is.EqualTo(-15d));
                Assert.That(wb.Evaluate("VALUE(\"(1,5e3)%\")"), Is.EqualTo(-15d));
            });

            var expectedSerialDate = new DateTime(2022, 3, 5).ToSerialDateTime();
            Assert.Multiple(() =>
            {
                Assert.That(wb.Evaluate("VALUE(\"5-březen-22\")"), Is.EqualTo(expectedSerialDate));
                Assert.That(wb.Evaluate("VALUE(\"05.03.2022\")"), Is.EqualTo(expectedSerialDate));
                Assert.That(wb.Evaluate("VALUE(\"5-březen\")"), Is.EqualTo(new DateTime(DateTime.Now.Year, 3, 5).ToSerialDateTime()));
            });
        }
    }
}
