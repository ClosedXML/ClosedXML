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
            Assert.AreEqual(0, SearchWildcard(text, string.Empty));
        }

        [TestCase("", "abc", 0)]
        [TestCase("a", "abc", 0)]
        [TestCase("ab", "abc", 0)]
        [TestCase("abc", "abc", 0)]
        [TestCase("bc", "abc", 1)]
        [TestCase("c", "abc", 2)]
        public void Substring_Of_Text_Matches_Text(string substringPattern, string text, int expectedIndex)
        {
            Assert.AreEqual(expectedIndex, SearchWildcard(text, substringPattern));
        }

        [TestCase("abcd", "abc")]
        public void Pattern_Not_In_Text_Returns_Negative_One(string pattern, string text)
        {
            Assert.AreEqual(-1, SearchWildcard(text, pattern));
        }

        [Test]
        public void Pattern_Comparison_Is_Case_Insensitive()
        {
            Assert.AreEqual(1, SearchWildcard("zabcd", "AbCd"));
        }

        [Test]
        public void Tilde_Is_Escape_Char()
        {
            Assert.AreEqual(1, SearchWildcard("_abc_", "~a~B~c"));
        }

        [TestCase("~*", "*", 0)]
        [TestCase("~*", "a", -1)]
        [TestCase("~?", "?", 0)]
        [TestCase("~?", "a", -1)]
        [TestCase("~a~b~", "ab", 0)]
        public void Escaped_Wildcards_Are_Matched_As_Chars(string pattern, string text, int expectedPosition)
        {
            Assert.AreEqual(expectedPosition, SearchWildcard(text, pattern));
        }

        [Test]
        public void Question_Mark_Wildcard_Matches_Any_Char()
        {
            Assert.AreEqual(0, SearchWildcard("abc", "a?c"));
        }

        [TestCase("abcd", "ab*cd", 0)]
        [TestCase(@"aaab_____cd", "ab*cd", 2)]
        [TestCase("*abc*", "***a*b*c***", 0)]

        public void Star_Wildcard_Matches_Any_Number_Of_Chars(string text, string pattern, int index)
        {
            Assert.AreEqual(index, SearchWildcard(text, pattern));
        }

        [Test]
        public void Unpaired_Escape_Char_At_The_End_Of_Pattern_Is_Not_Char()
        {
            Assert.AreEqual(0, SearchWildcard("a", "a~"));
        }

        [Test]
        public void Star_Wildcard_At_The_Beginning_Matches_First_Char()
        {
            Assert.AreEqual(0, SearchWildcard("abcccd", "*ccd"));
        }

        [Test]
        public void Pattern_Size_Is_Limited_To_255_Chars()
        {
            Assert.AreEqual(0, SearchWildcard(new string('a', 1000), new string('a', 255)));

            Assert.AreEqual(-1, SearchWildcard(new string('a', 1000), new string('a', 256)));
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
            Assert.AreEqual(matches, Wildcard.Matches(pattern.AsSpan(), text.AsSpan()));
        }

        private static int SearchWildcard(string text, string pattern)
        {
            return new Wildcard(pattern).Search(text.AsSpan());
        }
    }
}
