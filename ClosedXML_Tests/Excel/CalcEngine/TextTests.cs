using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Threading;

namespace ClosedXML_Tests.Excel.CalcEngine
{
    [TestFixture]
    public class TextTests
    {
        [OneTimeSetUp]
        public void Init()
        {
            // Make sure tests run on a deterministic culture
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        }

        [Test]
        public void Char_Empty_Input_String()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Char("""")"), Throws.Exception);
        }

        [Test]
        public void Char_Input_Too_Large()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Char(9797)"), Throws.Exception);
        }

        [Test]
        public void Char_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Char(97)");
            Assert.AreEqual("a", actual);
        }

        [Test]
        public void Clean_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Clean("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Clean_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Clean(CHAR(9)&""Monthly report""&CHAR(10))");
            Assert.AreEqual("Monthly report", actual);

            actual = XLWorkbook.EvaluateExpr(@"Clean(""   "")");
            Assert.AreEqual("   ", actual);
        }

        [Test]
        public void Code_Empty_Input_String()
        {
            // Todo: more specific exception - ValueException?
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Code("""")"), Throws.Exception);
        }

        [Test]
        public void Code_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Code(""A"")");
            Assert.AreEqual(65, actual);

            actual = XLWorkbook.EvaluateExpr(@"Code(""BCD"")");
            Assert.AreEqual(66, actual);
        }

        [Test]
        public void Concat_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Concatenate(""ABC"", ""123"")");
            Assert.AreEqual("ABC123", actual);

            actual = XLWorkbook.EvaluateExpr(@"Concatenate("""", ""123"")");
            Assert.AreEqual("123", actual);
        }

        [Test]
        public void Dollar_Empty_Input_String()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Dollar("", 3)"), Throws.Exception);
        }

        [Test]
        public void Dollar_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Dollar(123.54)");
            Assert.AreEqual("$123.54", actual);

            actual = XLWorkbook.EvaluateExpr(@"Dollar(123.54, 3)");
            Assert.AreEqual("$123.540", actual);
        }

        [Test]
        public void Exact_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Exact("""", """")");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void Exact_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""asdf"")");
            Assert.AreEqual(true, actual);

            actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""ASDF"")");
            Assert.AreEqual(false, actual);

            actual = XLWorkbook.EvaluateExpr(@"Exact(123, 123)");
            Assert.AreEqual(true, actual);

            actual = XLWorkbook.EvaluateExpr(@"Exact(321, 123)");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void Find_Start_Position_Too_Large()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Find(""abc"", ""abcdef"", 10)"), Throws.Exception);
        }

        [Test]
        public void Find_String_In_Another_Empty_String()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Find(""abc"", """")"), Throws.Exception);
        }

        [Test]
        public void Find_String_Not_Found()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Find(""123"", ""asdf"")"), Throws.Exception);
        }

        [Test]
        public void Find_Case_Sensitive_String_Not_Found()
        {
            // Find is case-sensitive
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Find(""excel"", ""Microsoft Excel 2010"")"), Throws.Exception);
        }

        [Test]
        public void Find_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Find(""Tuesday"", ""Today is Tuesday"")");
            Assert.AreEqual(10, actual);

            actual = XLWorkbook.EvaluateExpr(@"Find("""", """")");
            Assert.AreEqual(1, actual);

            actual = XLWorkbook.EvaluateExpr(@"Find("""", ""asdf"")");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Fixed_Input_Is_String()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Fixed(""asdf"")"), Throws.Exception);
        }

        [Test]
        public void Fixed_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Fixed(17300.67, 4)");
            Assert.AreEqual("17,300.6700", actual);

            actual = XLWorkbook.EvaluateExpr(@"Fixed(17300.67, 2, TRUE)");
            Assert.AreEqual("17300.67", actual);

            actual = XLWorkbook.EvaluateExpr(@"Fixed(17300.67)");
            Assert.AreEqual("17,300.67", actual);
        }

        [Test]
        public void Left_Bigger_Than_Length()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"", 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Left_Default()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"")");
            Assert.AreEqual("A", actual);
        }

        [Test]
        public void Left_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Left_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Left(""ABC"", 2)");
            Assert.AreEqual("AB", actual);
        }

        [Test]
        public void Len_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Len("""")");
            Assert.AreEqual(0, actual);
        }

        [Test]
        public void Len_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Len(""word"")");
            Assert.AreEqual(4, actual);
        }

        [Test]
        public void Lower_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Lower("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Lower_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Lower(""AbCdEfG"")");
            Assert.AreEqual("abcdefg", actual);
        }

        [Test]
        public void Mid_Bigger_Than_Length()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 1, 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Mid_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid("""", 1, 1)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Mid_Start_After()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 5, 5)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Mid_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Mid(""ABC"", 2, 2)");
            Assert.AreEqual("BC", actual);
        }

        [Test]
        public void Proper_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Proper("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Proper_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Proper(""my name is francois botha"")");
            Assert.AreEqual("My Name Is Francois Botha", actual);
        }

        [Test]
        public void Replace_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Replace("""", 1, 1, ""newtext"")");
            Assert.AreEqual("newtext", actual);
        }

        [Test]
        public void Replace_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Replace(""Here is some obsolete text to replace."", 14, 13, ""new text"")");
            Assert.AreEqual("Here is some new text to replace.", actual);
        }

        [Test]
        public void Rept_Empty_Input_Strings()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Rept("""", 3)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Rept_Start_Is_Negative()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Rept(""Francois"", -1)"), Throws.Exception);
        }

        [Test]
        public void Rept_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Rept(""Francois Botha,"", 3)");
            Assert.AreEqual("Francois Botha,Francois Botha,Francois Botha,", actual);

            actual = XLWorkbook.EvaluateExpr(@"Rept(""123"", 5/2)");
            Assert.AreEqual("123123", actual);

            actual = XLWorkbook.EvaluateExpr(@"Rept(""Francois"", 0)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Right_Bigger_Than_Length()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"", 5)");
            Assert.AreEqual("ABC", actual);
        }

        [Test]
        public void Right_Default()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"")");
            Assert.AreEqual("C", actual);
        }

        [Test]
        public void Right_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Right_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Right(""ABC"", 2)");
            Assert.AreEqual("BC", actual);
        }

        [Test]
        public void Search_No_Parameters_With_Values()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search("""", """")"), Throws.Exception);
        }

        [Test]
        public void Search_Empty_Search_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Search("""", ""asdf"")");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Search_Start_Position_Too_Large()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search(""abc"", ""abcdef"", 10)"), Throws.Exception);
        }

        [Test]
        public void Search_Empty_Input_String()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search(""abc"", """")"), Throws.Exception);
        }

        [Test]
        public void Search_String_Not_Found()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search(""123"", ""asdf"")"), Throws.Exception);
        }

        [Test]
        public void Search_Wildcard_String_Not_Found()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search(""soft?2010"", ""Microsoft Excel 2010"")"), Throws.Exception);
        }

        [Test]
        public void Search_Start_Position_Too_Large2()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Search(""text"", ""This is some text"", 15)"), Throws.Exception);
        }

        // http://www.excel-easy.com/examples/find-vs-search.html
        [Test]
        public void Search_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Search(""Tuesday"", ""Today is Tuesday"")");
            Assert.AreEqual(10, actual);

            // Find is case-INsensitive
            actual = XLWorkbook.EvaluateExpr(@"Search(""excel"", ""Microsoft Excel 2010"")");
            Assert.AreEqual(11, actual);

            actual = XLWorkbook.EvaluateExpr(@"Search(""soft*2010"", ""Microsoft Excel 2010"")");
            Assert.AreEqual(6, actual);

            actual = XLWorkbook.EvaluateExpr(@"Search(""Excel 20??"", ""Microsoft Excel 2010"")");
            Assert.AreEqual(11, actual);

            actual = XLWorkbook.EvaluateExpr(@"Search(""text"", ""This is some text"", 14)");
            Assert.AreEqual(14, actual);
        }

        [Test]
        public void Substitute_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Substitute(""This is a Tuesday."", ""Tuesday"", ""Monday"")");
            Assert.AreEqual("This is a Monday.", actual);

            actual = XLWorkbook.EvaluateExpr(@"Substitute(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 1)");
            Assert.AreEqual("This is a Monday. Next week also has a Tuesday.", actual);

            actual = XLWorkbook.EvaluateExpr(@"Substitute(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 2)");
            Assert.AreEqual("This is a Tuesday. Next week also has a Monday.", actual);

            actual = XLWorkbook.EvaluateExpr(@"Substitute("""", """", ""Monday"")");
            Assert.AreEqual("", actual);

            actual = XLWorkbook.EvaluateExpr(@"Substitute(""This is a Tuesday. Next week also has a Tuesday."", """", ""Monday"")");
            Assert.AreEqual("This is a Tuesday. Next week also has a Tuesday.", actual);

            actual = XLWorkbook.EvaluateExpr(@"Substitute(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", """")");
            Assert.AreEqual("This is a . Next week also has a .", actual);
        }

        [Test]
        public void T_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"T("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void T_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"T(""asdf"")");
            Assert.AreEqual("asdf", actual);

            actual = XLWorkbook.EvaluateExpr(@"T(Today())");
            Assert.AreEqual("", actual);

            actual = XLWorkbook.EvaluateExpr(@"T(TRUE)");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Text_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Text(1913415.93, """")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Text_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Text(Date(2010, 1, 1), ""yyyy-MM-dd"")");
            Assert.AreEqual("2010-01-01", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(1469.07, ""0,000,000.00"")");
            Assert.AreEqual("0,001,469.07", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(1913415.93, ""#,000.00"")");
            Assert.AreEqual("1,913,415.93", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(2800, ""$0.00"")");
            Assert.AreEqual("$2800.00", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(0.4, ""0%"")");
            Assert.AreEqual("40%", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(Date(2010, 1, 1), ""MMMM yyyy"")");
            Assert.AreEqual("January 2010", actual);

            actual = XLWorkbook.EvaluateExpr(@"Text(Date(2010, 1, 1), ""M/d/y"")");
            Assert.AreEqual("1/1/10", actual);
        }

        [Test]
        public void Text_String_Input()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"TEXT(""211x"", ""#00"")");
            Assert.AreEqual("211x", actual);
        }

        [Test]
        public void Trim_EmptyInput_Striong()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Trim("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Trim_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Trim("" some text with padding   "")");
            Assert.AreEqual("some text with padding", actual);
        }

        [Test]
        public void Upper_Empty_Input_String()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Upper("""")");
            Assert.AreEqual("", actual);
        }

        [Test]
        public void Upper_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Upper(""AbCdEfG"")");
            Assert.AreEqual("ABCDEFG", actual);
        }

        [Test]
        public void Value_Input_String_Is_Not_A_Number()
        {
            Assert.That(() => XLWorkbook.EvaluateExpr(@"Value(""asdf"")"), Throws.Exception);
        }

        [Test]
        public void Value_Value()
        {
            Object actual = XLWorkbook.EvaluateExpr(@"Value(""123.54"")");
            Assert.AreEqual(123.54, actual);

            actual = XLWorkbook.EvaluateExpr(@"Value(654.32)");
            Assert.AreEqual(654.32, actual);
        }
    }
}
