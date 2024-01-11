using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    /// <summary>
    /// Equal/NotEqual operators in custom filter are wildcard filters, *NOT* comparator filter.
    /// LessThan/EqualOrLessThan/EqualOrGreaterThan/GreaterThan are comparator filters.
    /// </summary>
    [TestFixture]
    public class CustomFilterTests
    {
        [Test]
        public void EqualOrLessThan_with_logical_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.EqualOrLessThan(true))
                .AddTrue(false, true)
                .AddFalse(Blank.Value, 1, "FALSE", "TRUE", XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrLessThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.EqualOrLessThan(1))
                .Add(0.9, true)
                .Add(1.1, false)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrLessThan_with_text_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.EqualOrLessThan("b"))
                .AddTrue("", "A", "b", "B")
                .AddFalse("C", Blank.Value, 1, false, XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrLessThan_with_error_compares_against_numeric_types_of_error()
        {
            new AutoFilterTester(f => f.EqualOrLessThan(XLError.CellReference))
                .AddTrue(XLError.NullValue, XLError.IncompatibleValue, XLError.CellReference)
                .AddFalse(XLError.NameNotRecognized, 1, "#NULL!", "Test", "", true, false, Blank.Value)
                .AssertVisibility();
        }

        [Test]
        public void LessThan_with_logical_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.LessThan(true))
                .AddTrue(false)
                .AddFalse(true, -1, Blank.Value, 1, "FALSE", "TRUE", XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void LessThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.LessThan(2))
                .Add(1.1, true)
                .Add(2, false)
                .AssertVisibility();
        }

        [Test]
        public void LessThan_with_text_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.LessThan("b"))
                .AddTrue("", "A")
                .AddFalse("b", "B", "C", Blank.Value, 1, false, XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void LessThan_with_error_compares_against_numeric_types_of_error()
        {
            new AutoFilterTester(f => f.LessThan(XLError.CellReference))
                .AddTrue(XLError.NullValue, XLError.IncompatibleValue)
                .AddFalse(XLError.CellReference, XLError.NameNotRecognized, 1, "#NULL!", "Test", "", true, false, Blank.Value)
                .AssertVisibility();
        }

        [Test]
        public void GreaterThan_with_logical_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.GreaterThan(false))
                .AddTrue(true)
                .AddFalse(false, -1, Blank.Value, 1, "FALSE", "TRUE", XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void GreaterThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.GreaterThan(0))
                .Add(0.1, true)
                .AddFalse(-0.1, -1)
                .AssertVisibility();
        }

        [Test]
        public void GreaterThan_with_text_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.GreaterThan("b"))
                .AddTrue("C", "c")
                .AddFalse("", "A", "b", "B", Blank.Value, 1, false, XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void GreaterThan_with_error_compares_against_numeric_types_of_error()
        {
            new AutoFilterTester(f => f.GreaterThan(XLError.CellReference))
                .AddTrue(XLError.NameNotRecognized, XLError.NumberInvalid, XLError.NoValueAvailable)
                .AddFalse(XLError.CellReference, XLError.IncompatibleValue, XLError.NullValue, 1, "#NULL!", "Test", "", true, false, Blank.Value)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrGreaterThan_with_logical_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.EqualOrGreaterThan(false))
                .AddTrue(false, true)
                .AddFalse(-1, 0, Blank.Value, 1, "FALSE", "TRUE", XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrGreaterThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.EqualOrGreaterThan(1))
                .Add(0.9, false)
                .Add(1.1, true)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrGreaterThan_with_text_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.EqualOrGreaterThan("b"))
                .AddTrue("b", "B", "Ba", "C", "c")
                .AddFalse("", "A", Blank.Value, 1, false, XLError.NullValue)
                .AssertVisibility();
        }

        [Test]
        public void EqualOrGreaterThan_with_error_compares_against_numeric_types_of_error()
        {
            new AutoFilterTester(f => f.EqualOrGreaterThan(XLError.CellReference))
                .AddTrue(XLError.CellReference, XLError.NameNotRecognized, XLError.NumberInvalid, XLError.NoValueAvailable)
                .AddFalse(XLError.IncompatibleValue, XLError.NullValue, 1, "#NULL!", "Test", "", true, false, Blank.Value)
                .AssertVisibility();
        }

        [Test]
        public void Equal_uses_wildcard_matching_for_patterns_against_text_only()
        {
            new AutoFilterTester(f => f.EqualTo("1*0"))
                .AddTrue("1.0", "1 and 0")
                .AddFalse(1, "A", "B", 2, XLError.DivisionByZero, true, false)
                .Add(1, nf => nf.SetFormat("1.0"), false)
                .Add(1, nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2), false)
                .AssertVisibility();
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void Equal_uses_format_string_matching_for_filter_values_that_look_like_non_patterns()
        {
            // Note the ',' separator that is used detect number. Excel doesn't use invariant culture.
            new AutoFilterTester(f => f.EqualTo("1,00"))
                .Add("1,00", true)
                .Add(1, nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2), true)
                .Add(99, nf => nf.SetFormat("\"1,00\""), true)
                .AddFalse(1, "A", "B", 2, XLError.DivisionByZero, true, false)
                .AssertVisibility();
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void NotEqual_matches_detected_type_and_value_of_filter_value_for_non_text_data_types()
        {
            // 1,00 is detected as a type number with value 1.
            new AutoFilterTester(f => f.NotEqualTo("1,00"))
                .Add(1, false) // Value is equal => hide
                .Add(1, nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2), false) // Value is equal => hide
                .Add("1,00", true) // wrong type
                .Add(99, nf => nf.SetFormat("\"1,00\""), true) // Value is wrong => non-equal
                .AddTrue("A", "B", 2, XLError.DivisionByZero, true, false) // Wrong type
                .AssertVisibility();
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void NotEqual_for_detected_wildcard_matches_only_texts()
        {
            // NotEqual with text pattern must have text type.
            new AutoFilterTester(f => f.NotEqualTo("1*0"))
                .Add(1, true)
                .Add(1, nf => nf.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2), true)
                .Add("1,00", false)
                .Add("100", false)
                .Add(100, true)
                .Add(99, nf => nf.SetFormat("\"1,00\""), true)
                .AddTrue("A", "B", 2, XLError.DivisionByZero, true, false)
                .AssertVisibility();
        }

        private static AutoFilterTester WithOneAndOtherTypes(Action<IXLFilterColumn> filter)
        {
            // Add equivalent of 1 and other types
            return new AutoFilterTester(filter)
                .Add(1, true)
                .Add(new DateTime(1900, 1, 1), true) // =1 in serial date time
                .Add(new TimeSpan(1, 0, 0, 0), true) // =1 in serial date time
                .Add("1", false)
                .Add(Blank.Value, false)
                .Add("Hello", false)
                .Add(true, false)
                .Add(XLError.NullValue, false); // #NULL! has type value 1
        }
    }
}
