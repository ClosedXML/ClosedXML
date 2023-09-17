using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class XLCalculationChainTests
    {
        [Test]
        public void Enumerating_empty_chain()
        {
            var chain = new XLCalculationChain();
            CollectionAssert.IsEmpty(GetList(chain));
        }

        [Test]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [TestCase(40)]
        public void Enumerating_whole_chain(int chainLength)
        {
            var chain = new XLCalculationChain();
            var expectedPoints = new List<XLBookPoint>();
            for (var i = 0; i < chainLength; ++i)
            {
                var point = new XLBookPoint(1, new XLSheetPoint(1, i));
                chain.AddLast(point);
                expectedPoints.Add(point);
            }

            CollectionAssert.AreEqual(expectedPoints, GetList(chain));
        }

        [Test]
        public void Enumeration_ends_on_target()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));

            chain.AddLast(a1);
            chain.AddLast(b1);
            chain.AddLast(c1);

            CollectionAssert.AreEqual(new[] { a1, b1 }, GetList(chain, b1));
        }

        [Test]
        public void Remove_throws_on_missing_point()
        {
            var chain = new XLCalculationChain();

            Assert.Throws<InvalidOperationException>(
                () => chain.Remove(new XLBookPoint(1, new XLSheetPoint(1, 1))));
        }

        [Test]
        public void Remove_link_from_chain()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            var d1 = new XLBookPoint(1, new XLSheetPoint(1, 4));

            chain.AddLast(a1);
            chain.AddLast(b1);
            chain.AddLast(c1);
            chain.AddLast(d1);

            // Remove point in the middle
            chain.Remove(c1);
            CollectionAssert.AreEqual(new[] { a1, b1, d1 }, GetList(chain));

            // Remove last point in the sequence
            chain.Remove(d1);
            CollectionAssert.AreEqual(new[] { a1, b1 }, GetList(chain));

            // Remove head
            chain.Remove(a1);
            CollectionAssert.AreEqual(new[] { b1 }, GetList(chain));

            // Remove the only remaining
            chain.Remove(b1);
            CollectionAssert.IsEmpty(GetList(chain));
        }

        [Test]
        public void AddAfter_adds_point()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            chain.AddLast(a1);

            // Add as tail for single link chain
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            chain.AddAfter(a1, b1);
            CollectionAssert.AreEqual(new[] { a1, b1 }, GetList(chain));

            // Add as tail for multi link chain
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddAfter(b1, c1);
            CollectionAssert.AreEqual(new[] { a1, b1, c1 }, GetList(chain));

            // Add somewhere in the middle
            var d1 = new XLBookPoint(1, new XLSheetPoint(1, 4));
            chain.AddAfter(b1, d1);
            CollectionAssert.AreEqual(new[] { a1, b1, d1, c1 }, GetList(chain));
        }

        [Test]
        public void MoveToFront_moves_the_point_to_the_front()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            chain.AddLast(a1);
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            chain.AddLast(b1);
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddLast(c1);
            var d1 = new XLBookPoint(1, new XLSheetPoint(1, 4));
            chain.AddLast(d1);

            var enumerator = chain.GetEnumerator(null);
            Assert.True(enumerator.MoveAhead());
            Assert.AreEqual(a1, enumerator.Point);

            // a,b,c,d -> d,a,b,c
            enumerator.MoveToFront(d1);
            Assert.AreEqual(d1, enumerator.Point);
            Assert.AreEqual(new[] { d1, a1, b1, c1 }, GetList(chain));

            // d,a,b,c -> b,d,a,c
            enumerator.MoveToFront(b1);
            Assert.AreEqual(b1, enumerator.Point);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, GetList(chain));

            Assert.True(enumerator.MoveAhead());
            Assert.AreEqual(d1, enumerator.Point);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, GetList(chain));

            // d,a,c -> a,d,c
            enumerator.MoveToFront(a1);
            Assert.AreEqual(a1, enumerator.Point);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, GetList(chain));

            // Move A1 to front when it's already at front
            enumerator.MoveToFront(a1);
            Assert.AreEqual(a1, enumerator.Point);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, GetList(chain));

            // a,d,c -> c,a,d
            enumerator.MoveToFront(c1);
            Assert.AreEqual(c1, enumerator.Point);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, GetList(chain));

            Assert.True(enumerator.MoveAhead());
            Assert.AreEqual(a1, enumerator.Point);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, GetList(chain));

            // a,d -> d,a
            enumerator.MoveToFront(d1);
            Assert.AreEqual(d1, enumerator.Point);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetList(chain));

            Assert.True(enumerator.MoveAhead());
            Assert.AreEqual(a1, enumerator.Point);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetList(chain));

            // a -> a
            enumerator.MoveToFront(a1);
            Assert.AreEqual(a1, enumerator.Point);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetList(chain));

            Assert.False(enumerator.MoveAhead());
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetList(chain));
        }

        private static IEnumerable<XLBookPoint> GetList(XLCalculationChain chain, XLBookPoint? target = null)
        {
            var enumerator = chain.GetEnumerator(target);
            while (enumerator.MoveAhead())
            {
                yield return enumerator.Point;
            }
        }
    }
}
