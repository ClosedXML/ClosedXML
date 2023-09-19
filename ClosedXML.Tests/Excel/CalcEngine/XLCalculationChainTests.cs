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
            CollectionAssert.IsEmpty(chain.GetPoints());
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

            CollectionAssert.AreEqual(expectedPoints, chain.GetPoints());
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
            CollectionAssert.AreEqual(new[] { a1, b1, d1 }, chain.GetPoints());

            // Remove last point in the sequence
            chain.Remove(d1);
            CollectionAssert.AreEqual(new[] { a1, b1 }, chain.GetPoints());

            // Remove head
            chain.Remove(a1);
            CollectionAssert.AreEqual(new[] { b1 }, chain.GetPoints());

            // Remove the only remaining
            chain.Remove(b1);
            CollectionAssert.IsEmpty(chain.GetPoints());
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
            CollectionAssert.AreEqual(new[] { a1, b1 }, chain.GetPoints());

            // Add as tail for multi link chain
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddAfter(b1, c1);
            CollectionAssert.AreEqual(new[] { a1, b1, c1 }, chain.GetPoints());

            // Add somewhere in the middle
            var d1 = new XLBookPoint(1, new XLSheetPoint(1, 4));
            chain.AddAfter(b1, d1);
            CollectionAssert.AreEqual(new[] { a1, b1, d1, c1 }, chain.GetPoints());
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

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(a1, chain.Current);

            // a,b,c,d -> d,a,b,c
            chain.MoveToCurrent(d1);
            Assert.AreEqual(d1, chain.Current);
            Assert.AreEqual(new[] { d1, a1, b1, c1 }, chain.GetPoints());

            // d,a,b,c -> b,d,a,c
            chain.MoveToCurrent(b1);
            Assert.AreEqual(b1, chain.Current);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, chain.GetPoints());

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(d1, chain.Current);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, chain.GetPoints());

            // d,a,c -> a,d,c
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, chain.GetPoints());

            // Move A1 to front when it's already at front
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, chain.GetPoints());

            // a,d,c -> c,a,d
            chain.MoveToCurrent(c1);
            Assert.AreEqual(c1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, chain.GetPoints());

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, chain.GetPoints());

            // a,d -> d,a
            chain.MoveToCurrent(d1);
            Assert.AreEqual(d1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, chain.GetPoints());

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, chain.GetPoints());

            // a -> a
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, chain.GetPoints());

            Assert.False(chain.MoveAhead());
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, chain.GetPoints());
        }
    }
}
