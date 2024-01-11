using System;
using System.Collections.Generic;
using System.Linq;
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
            CollectionAssert.IsEmpty(GetPoints(chain));
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

            CollectionAssert.AreEqual(expectedPoints, GetPoints(chain));
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
            CollectionAssert.AreEqual(new[] { a1, b1, d1 }, GetPoints(chain));

            // Remove last point in the sequence
            chain.Remove(d1);
            CollectionAssert.AreEqual(new[] { a1, b1 }, GetPoints(chain));

            // Remove head
            chain.Remove(a1);
            CollectionAssert.AreEqual(new[] { b1 }, GetPoints(chain));

            // Remove the only remaining
            chain.Remove(b1);
            CollectionAssert.IsEmpty(GetPoints(chain));
        }

        [Test]
        public void AddAfter_adds_point()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            chain.AddLast(a1);

            // Add as tail for single link chain
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            chain.AddAfter(a1, b1, 0);
            CollectionAssert.AreEqual(new[] { a1, b1 }, GetPoints(chain));

            // Add as tail for multi link chain
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddAfter(b1, c1, 0);
            CollectionAssert.AreEqual(new[] { a1, b1, c1 }, GetPoints(chain));

            // Add somewhere in the middle
            var d1 = new XLBookPoint(1, new XLSheetPoint(1, 4));
            chain.AddAfter(b1, d1, 0);
            CollectionAssert.AreEqual(new[] { a1, b1, d1, c1 }, GetPoints(chain));
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
            Assert.AreEqual(new[] { d1, a1, b1, c1 }, GetPoints(chain));

            // d,a,b,c -> b,d,a,c
            chain.MoveToCurrent(b1);
            Assert.AreEqual(b1, chain.Current);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, GetPoints(chain));

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(d1, chain.Current);
            Assert.AreEqual(new[] { b1, d1, a1, c1 }, GetPoints(chain));

            // d,a,c -> a,d,c
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, GetPoints(chain));

            // Move A1 to front when it's already at front
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, a1, d1, c1 }, GetPoints(chain));

            // a,d,c -> c,a,d
            chain.MoveToCurrent(c1);
            Assert.AreEqual(c1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, GetPoints(chain));

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, a1, d1 }, GetPoints(chain));

            // a,d -> d,a
            chain.MoveToCurrent(d1);
            Assert.AreEqual(d1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

            Assert.True(chain.MoveAhead());
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

            // a -> a
            chain.MoveToCurrent(a1);
            Assert.AreEqual(a1, chain.Current);
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

            Assert.False(chain.MoveAhead());
            Assert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));
        }

        [Test]
        public void Traversal_detects_cycles()
        {
            var chain = new XLCalculationChain();
            // `=C1+B1`
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            chain.AddLast(a1);
            // `=A1`
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            chain.AddLast(b1);
            // `=A1`
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddLast(c1);

            // Move to the first link.
            Assert.True(chain.MoveAhead());

            // Cycle a1, c1, when we first encounter c1, we don't know yet that it's a cycle
            chain.MoveToCurrent(c1);
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));

            // A1 is marked with a position, because they have been at the current
            // C1 hasn't ben pushed back yet, so it keeps 0.
            CollectionAssert.AreEqual(new[] { 0, 1, 0 }, GetPositions(chain));

            // But then we get A1 again, without any other point being marked
            // as done, therefore we are at cycle.
            chain.MoveToCurrent(a1);
            CollectionAssert.AreEqual(new[] { a1, c1, b1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));
            Assert.True(chain.IsCurrentInCycle);

            // When we encounter C1 again, it's obviously a cycle.
            chain.MoveToCurrent(c1);
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));
            Assert.True(chain.IsCurrentInCycle);

            // Let's move on and get A1 to the current. Because the C1 has been
            // marked as done, A1 is no longer in cycle.
            chain.MoveAhead();
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));

            // C1 position has been cleared, because it has moved beyond
            // current and A1 is now current.
            CollectionAssert.AreEqual(new[] { 0, 1, 0 }, GetPositions(chain));

            // A1 is no longer in a current, because current position is 2, but last position
            // of A1 was 1 => there has been a processed node in the meantime.
            Assert.False(chain.IsCurrentInCycle);

            chain.MoveToCurrent(b1);
            CollectionAssert.AreEqual(new[] { c1, b1, a1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 0, 0, 2 }, GetPositions(chain));
            Assert.False(chain.IsCurrentInCycle);

            chain.MoveToCurrent(a1);
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 0, 2, 2 }, GetPositions(chain));
            Assert.True(chain.IsCurrentInCycle);

            chain.MoveAhead();
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 0, 0, 2 }, GetPositions(chain));
            Assert.False(chain.IsCurrentInCycle);

            chain.MoveAhead();
            CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 0, 0, 0 }, GetPositions(chain));
        }

        [Test]
        public void Reset_clears_positions_ahead_of_current()
        {
            var chain = new XLCalculationChain();
            var a1 = new XLBookPoint(1, new XLSheetPoint(1, 1));
            chain.AddLast(a1);
            var b1 = new XLBookPoint(1, new XLSheetPoint(1, 2));
            chain.AddLast(b1);
            var c1 = new XLBookPoint(1, new XLSheetPoint(1, 3));
            chain.AddLast(c1);

            Assert.True(chain.MoveAhead());

            chain.MoveToCurrent(b1);
            chain.MoveToCurrent(a1);
            Assert.True(chain.IsCurrentInCycle);
            CollectionAssert.AreEqual(new[] { a1, b1, c1 }, GetPoints(chain));
            CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));

            chain.Reset();

            CollectionAssert.AreEqual(new[] { 0, 0, 0 }, GetPositions(chain));
        }

        private static IEnumerable<XLBookPoint> GetPoints(XLCalculationChain chain)
        {
            return chain.GetLinks().Select(x => x.Point);
        }

        private static IEnumerable<int> GetPositions(XLCalculationChain chain)
        {
            return chain.GetLinks().Select(x => x.LastPosition);
        }
    }
}
