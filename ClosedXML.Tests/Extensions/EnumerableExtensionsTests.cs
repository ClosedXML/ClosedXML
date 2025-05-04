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
            var array = Array.Empty<int>();
            Assert.That(array.GetItemType(), Is.EqualTo(typeof(int)));

            var list = new List<double>();
            Assert.Multiple(() =>
            {
                Assert.That(list.GetItemType(), Is.EqualTo(typeof(double)));
                Assert.That(list.AsEnumerable().GetItemType(), Is.EqualTo(typeof(double)));
            });

            IEnumerable<IEnumerable> enumerable = new List<string>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(string)));

            enumerable = new List<List<string>>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(List<string>)));

            enumerable = new List<int[]>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(int[])));

            var anonymousIterator = new List<TablesTests.TestObjectWithoutAttributes>()
                .Select(o => new { FirstName = o.Column1, LastName = o.Column2 });

            //expectedType can be something like <>f__AnonymousType9`2[System.String,System.String]
            //but since that `9` may differ with new anonymous types declare in the assembly
            //check the beginning and the ending of the actual type
            var expectedTypeStart = "<>f__AnonymousType";
            var expectedTypeEnd = "`2[System.String,System.String]";
            var actualType = anonymousIterator.GetItemType().ToString();
            Assert.Multiple(() =>
            {
                Assert.That(actualType, Does.StartWith(expectedTypeStart));
                Assert.That(actualType, Does.EndWith(expectedTypeEnd));
            });

            IEnumerable<object> obj = anonymousIterator;
            actualType = obj.GetItemType().ToString();
            Assert.Multiple(() =>
            {
                Assert.That(actualType, Does.StartWith(expectedTypeStart));
                Assert.That(actualType, Does.EndWith(expectedTypeEnd));
            });
        }

        [Test]
        public void SkipLast_skips_last_element_of_enumerable()
        {
            var empty = Array.Empty<int>().SkipLast();
            Assert.That(empty, Is.Empty);

            var oneElement = new[] { 1 }.SkipLast();
            Assert.That(oneElement, Is.Empty);

            var twoElements = new[] { 1, 2 }.SkipLast();
            Assert.That(twoElements, Is.EqualTo(new[] { 1 }).AsCollection);
        }

        [Test]
        public void WhereNotNull_removes_null_elements()
        {
            var source = new int?[] { 1, null, 2 };

            var result = source.WhereNotNull(x => x);

            Assert.That(result, Is.EqualTo(new[] { 1, 2 }).AsCollection);
        }
    }
}
