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
        public void EqualLessThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.EqualOrLessThan(1))
                .Add(0.9, true)
                .Add(1.1, false)
                .AssertVisibility();
        }

        [Test]
        public void EqualLessThan_with_text_compares_against_values_of_same_type()
        {
            new AutoFilterTester(f => f.EqualOrLessThan("b"))
                .Add("", true).Add("A", true).Add("b", true).Add("B", true).Add("C", false)
                .Add(Blank.Value, false).Add(1, false).Add(false, false).Add(XLError.NullValue, false)
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
        public void GreaterThan_with_number_compares_against_values_of_same_type()
        {
            WithOneAndOtherTypes(f => f.GreaterThan(0))
                .Add(0.1, true)
                .Add(-0.1, false)
                .Add(-1, false)
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

        private AutoFilterTester WithOneAndOtherTypes(Action<IXLFilterColumn> filter)
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
