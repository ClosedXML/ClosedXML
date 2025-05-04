using System;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class WildcardTests
    {
        [TestCase("")]
        [TestCase("abc")]
        public void Empty_Pattern_Matches_Any_String(string text)
        {
            Assert.That(SearchWildcard(text, string.Empty), Is.EqualTo(0));
        }

        [TestCase("", "abc", 0)]
        [TestCase("a", "abc", 0)]
        [TestCase("ab", "abc", 0)]
        [TestCase("abc", "abc", 0)]
        [TestCase("bc", "abc", 1)]
        [TestCase("c", "abc", 2)]
        public void Substring_Of_Text_Matches_Text(string substringPattern, string text, int expectedIndex)
        {
            Assert.That(SearchWildcard(text, substringPattern), Is.EqualTo(expectedIndex));
        }

        [TestCase("abcd", "abc")]
        public void Pattern_Not_In_Text_Returns_Negative_One(string pattern, string text)
        {
            Assert.That(SearchWildcard(text, pattern), Is.EqualTo(-1));
        }

        [Test]
        public void Pattern_Comparison_Is_Case_Insensitive()
        {
            Assert.That(SearchWildcard("zabcd", "AbCd"), Is.EqualTo(1));
        }

        [Test]
        public void Tilde_Is_Escape_Char()
        {
            Assert.That(SearchWildcard("_abc_", "~a~B~c"), Is.EqualTo(1));
        }

        [TestCase("~*", "*", 0)]
        [TestCase("~*", "a", -1)]
        [TestCase("~?", "?", 0)]
        [TestCase("~?", "a", -1)]
        [TestCase("~a~b~", "ab", 0)]
        public void Escaped_Wildcards_Are_Matched_As_Chars(string pattern, string text, int expectedPosition)
        {
            Assert.That(SearchWildcard(text, pattern), Is.EqualTo(expectedPosition));
        }

        [Test]
        public void Question_Mark_Wildcard_Matches_Any_Char()
        {
            Assert.That(SearchWildcard("abc", "a?c"), Is.EqualTo(0));
        }

        [TestCase("abcd", "ab*cd", 0)]
        [TestCase(@"aaab_____cd", "ab*cd", 2)]
        [TestCase("*abc*", "***a*b*c***", 0)]

        public void Star_Wildcard_Matches_Any_Number_Of_Chars(string text, string pattern, int index)
        {
            Assert.That(SearchWildcard(text, pattern), Is.EqualTo(index));
        }

        [Test]
        public void Unpaired_Escape_Char_At_The_End_Of_Pattern_Is_Not_Char()
        {
            Assert.That(SearchWildcard("a", "a~"), Is.EqualTo(0));
        }

        [Test]
        public void Star_Wildcard_At_The_Beginning_Matches_First_Char()
        {
            Assert.That(SearchWildcard("abcccd", "*ccd"), Is.EqualTo(0));
        }

        [Test]
        public void Pattern_Size_Is_Limited_To_255_Chars()
        {
            Assert.Multiple(() =>
            {
                Assert.That(SearchWildcard(new string('a', 1000), new string('a', 255)), Is.EqualTo(0));

                Assert.That(SearchWildcard(new string('a', 1000), new string('a', 256)), Is.EqualTo(-1));
            });
        }

        [TestCase("?", "a", true)]
        [TestCase("?", "ab", false)]
        [TestCase("a?", "ab", true)]
        [TestCase("a?", "abc", false)]
        [TestCase("?b", "ab", true)]
        [TestCase("?b", "aab", false)]
        [TestCase("a*", "abc", true)]
        [TestCase("*a*", "abc", true)]
        [TestCase("*c", "abc", true)]
        [TestCase("*a*a", "abc", false)]
        [TestCase("*a*a", "aba", true)]
        [TestCase("*a*a", @"zaba", true)]
        [TestCase("a*", @"zaba", false)]
        public void Matches(string pattern, string text, bool matches)
        {
            Assert.That(new Wildcard(pattern).Matches(text.AsSpan()), Is.EqualTo(matches));
        }

        private static int SearchWildcard(string text, string pattern)
        {
            return new Wildcard(pattern).Search(text.AsSpan());
        }
    }
}
