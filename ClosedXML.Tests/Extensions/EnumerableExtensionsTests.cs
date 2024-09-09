using ClosedXML.Excel;
using ClosedXML.Tests.Excel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Extensions
{
    public class EnumerableExtensionsTests
    {
        [Test]
        public void CanGetItemType()
        {
            var array = new int[0];
            Assert.AreEqual(typeof(int), array.GetItemType());

            var list = new List<double>();
            Assert.AreEqual(typeof(double), list.GetItemType());
            Assert.AreEqual(typeof(double), list.AsEnumerable().GetItemType());

            IEnumerable<IEnumerable> enumerable = new List<string>();
            Assert.AreEqual(typeof(string), enumerable.GetItemType());

            enumerable = new List<List<string>>();
            Assert.AreEqual(typeof(List<string>), enumerable.GetItemType());

            enumerable = new List<int[]>();
            Assert.AreEqual(typeof(int[]), enumerable.GetItemType());

            var anonymousIterator = new List<TablesTests.TestObjectWithoutAttributes>()
                .Select(o => new { FirstName = o.Column1, LastName = o.Column2 });

            //expectedType can be something like <>f__AnonymousType9`2[System.String,System.String]
            //but since that `9` may differ with new anonymous types declare in the assembly
            //check the beginning and the ending of the actual type
            var expectedTypeStart = "<>f__AnonymousType";
            var expectedTypeEnd = "`2[System.String,System.String]";
            var actualType = anonymousIterator.GetItemType().ToString();
            Assert.True(actualType.StartsWith(expectedTypeStart));
            Assert.True(actualType.EndsWith(expectedTypeEnd));

            IEnumerable<object> obj = anonymousIterator;
            actualType = obj.GetItemType().ToString();
            Assert.True(actualType.StartsWith(expectedTypeStart));
            Assert.True(actualType.EndsWith(expectedTypeEnd));
        }

        [Test]
        public void SkipLast_skips_last_element_of_enumerable()
        {
            var empty = Array.Empty<int>().SkipLast();
            CollectionAssert.IsEmpty(empty);

            var oneElement = new[] { 1 }.SkipLast();
            CollectionAssert.IsEmpty(oneElement);

            var twoElements = new[] { 1, 2 }.SkipLast();
            CollectionAssert.AreEqual(new[] { 1 }, twoElements);
        }

        [Test]
        public void WhereNotNull_removes_null_elements()
        {
            var source = new int?[] { 1, null, 2 };

            var result = source.WhereNotNull(x => x);

            CollectionAssert.AreEqual(new[] { 1, 2 }, result);
        }
    }
}
