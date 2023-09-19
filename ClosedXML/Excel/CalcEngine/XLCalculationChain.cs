using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// <para>
    /// A calculation chain of formulas. Contains all formulas in the workbook.
    /// </para>
    /// <para>
    /// Calculation chain is an ordering of all cells that have value calculated
    /// by a formula (note that one formula can determine value of multiple cells,
    /// e.g. array). Formulas are calculated in specified order and if currently
    /// processed formula needs data from a cell whose value is dirty (i.e. it
    /// is determined by a not-yet-calculated formula), the current formula is
    /// stopped and the required formula is placed before the current one and starts
    /// to be processed. Once it is done, the original formula is starts to be processed
    /// again. It might have encounter another not-yet-calculated formula or it
    /// will finish and the calculation chain moves to the next one.
    /// </para>
    /// <para>
    /// Chain can be traversed through <see cref="Current"/>, <see cref="MoveAhead"/>,
    /// <see cref="MoveToCurrent"/> and <see cref="Reset"/>, but only one enumeration
    /// can go on at the time due to cycle detection.
    /// </para>
    /// </summary>
    internal class XLCalculationChain
    {
        /// <summary>
        /// Key to the <see cref="_nodeMap"/> that is the head of the chain.
        /// Null, when chain is empty.
        /// </summary>
        private XLBookPoint? _head;

        /// <summary>
        /// Key to the <see cref="_nodeMap"/> that is the tail of the chain.
        /// Null, when chain is empty.
        /// </summary>
        private XLBookPoint? _tail;

        /// <summary>
        /// <para>
        /// Doubly circular linked list containing all points with value
        /// calculated by a formula. The chain is "looped", so it doesn't
        /// have to deal with nulls for <see cref="XLBookPoint"/>.
        /// </para>
        /// <para>
        /// There is always exactly one loop, no cycles. The formulas might
        /// cause cycles due to dependencies, but that is manifested by
        /// constantly switching the links in a loop.</para>
        /// </summary>
        private readonly Dictionary<XLBookPoint, Link> _nodeMap = new();

        private XLBookPoint? _current;

        /// <summary>
        /// The address of a current of the chain.
        /// </summary>
        internal XLBookPoint Current => _current!.Value;

        /// <summary>
        /// Add a new link at the beginning of a chain.
        /// </summary>
        internal void AddFirst(XLBookPoint point)
        {
            if (_head is null || _tail is null)
            {
                Init(point);
                return;
            }

            Insert(point, _tail.Value, _head.Value);
            _head = point;
        }

        /// <summary>
        /// Append formula at the end of the chain.
        /// </summary>
        internal void AddLast(XLBookPoint point)
        {
            if (_head is null || _tail is null)
            {
                Init(point);
                return;
            }

            Insert(point, _tail.Value, _head.Value);
            _tail = point;
        }

        /// <summary>
        /// Initialize empty chain with a single link chain.
        /// </summary>
        private void Init(XLBookPoint point)
        {
            Debug.Assert(_nodeMap.Count == 0 && _head is null && _tail is null);
            _nodeMap.Add(point, new Link(point, point));
            _head = _tail = point;
        }

        /// <summary>
        /// Insert a link into the <see cref="_nodeMap"/> between
        /// <paramref name="prev"/> and <paramref name="next"/>.
        /// Don't update head or tail.
        /// </summary>
        private void Insert(XLBookPoint point, XLBookPoint prev, XLBookPoint next)
        {
            _nodeMap.Add(point, new Link(prev, next));

            var prevLink = _nodeMap[prev];
            _nodeMap[prev] = new Link(prevLink.Previous, point);

            var nextLink = _nodeMap[next];
            _nodeMap[next] = new Link(point, nextLink.Next);
        }

        /// <summary>
        /// Add a link for <paramref name="point"/> after the link for
        /// <paramref name="anchor"/>.
        /// </summary>
        /// <param name="anchor">
        /// The anchor point after which will be the new point added.
        /// </param>
        /// <param name="point">Point to add to the chain.</param>
        internal void AddAfter(XLBookPoint anchor, XLBookPoint point)
        {
            var prevLink = _nodeMap[anchor];
            var next = prevLink.Next;
            Insert(point, anchor, next);

            if (anchor == _tail!.Value)
                _tail = point;
        }

        /// <summary>
        /// Remove point from the chain.
        /// </summary>
        /// <param name="point">Point to remove.</param>
        /// <exception cref="InvalidOperationException">Point is not a part of the chain.</exception>
        internal void Remove(XLBookPoint point)
        {
            if (!_nodeMap.TryGetValue(point, out var pointLink))
                throw PointNotInChain(point);

            // Point is in the chain and there is exactly one link -> clear all.
            if (_nodeMap.Count == 1)
            {
                Clear();
                return;
            }

            if (point == _head!.Value)
                _head = pointLink.Next;

            if (point == _tail!.Value)
                _tail = pointLink.Previous;

            var prevLink = _nodeMap[pointLink.Previous];
            Debug.Assert(prevLink.Next == point);
            _nodeMap[pointLink.Previous] = new Link(prevLink.Previous, pointLink.Next);

            var nextLink = _nodeMap[pointLink.Next];
            Debug.Assert(nextLink.Previous == point);
            _nodeMap[pointLink.Next] = new Link(pointLink.Previous, nextLink.Next);

            _nodeMap.Remove(point);
        }

        /// <summary>
        /// Clear whole chain.
        /// </summary>
        internal void Clear()
        {
            _nodeMap.Clear();
            _head = null;
            _tail = null;
        }

        /// <summary>
        /// Enumerate all points in the chain.
        /// </summary>
        /// <returns></returns>
        internal IEnumerable<XLBookPoint> GetPoints()
        {
            if (_head is null)
                yield break;

            var current = _head.Value;
            do
            {
                yield return current;
                current = _nodeMap[current].Next;
            } while (current != _head.Value);
        }

        internal void Reset()
        {
            _current = null;
        }

        /// <summary>
        /// Mark current link as complete and move ahead to the next one.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the enumerator moved ahead, <c>false</c> if
        /// there are no more links and chain has moved passed target or
        /// looped completely.
        /// </returns>
        internal bool MoveAhead()
        {
            // First move
            if (_current is null)
            {
                var isChainEmpty = _head is null;
                if (isChainEmpty)
                    return false;

                _current = _head;
                return true;
            }

            // Subsequent move
            var currentPoint = _current.Value;
            if (!_nodeMap.TryGetValue(currentPoint, out var currentLink))
                throw PointNotInChain(currentPoint);

            var nextPoint = currentLink.Next;
            Debug.Assert(_nodeMap[nextPoint].Previous == currentPoint);
            if (nextPoint == _head!.Value)
            {
                // Whole chain has been calculated.
                return false;
            }

            _current = nextPoint;
            return true;
        }

        /// <summary>
        /// Move the <paramref name="pointToMove"/> before the current point
        /// as the new current to be calculated.
        /// </summary>
        /// <param name="pointToMove">
        /// The point of a chain to moved to the current. Should always be in
        /// the chain after the current.
        /// </param>
        internal void MoveToCurrent(XLBookPoint pointToMove)
        {
            if (_current is null)
                throw new InvalidOperationException("Enumerator not at a link.");

            var currentPoint = _current.Value;

            // If we are not moving anything, adding and removing doesn't
            // change chain, plus we avoid problems with moving in a
            // single/double link chain.
            if (currentPoint == pointToMove)
                return;

            // If head is also current, moving before the current means moving before head
            if (_head == currentPoint)
            {
                Remove(pointToMove);
                AddFirst(pointToMove);
            }
            else
            {
                // Current is not a head = move a link after prev of current.
                Remove(pointToMove);
                var anchor = _nodeMap[currentPoint].Previous;
                AddAfter(anchor, pointToMove);
            }

            _current = pointToMove;
        }

        private InvalidOperationException PointNotInChain(XLBookPoint point)
        {
            var exception = new InvalidOperationException($"Book point {point} is not in the chain.");
            exception.Data.Add("Chain", string.Join(", ", _nodeMap.Select(n => $"{n.Key}(prev:{n.Value.Previous},next:{n.Value.Next})")));
            return exception;
        }

        private readonly struct Link
        {
            internal readonly XLBookPoint Previous;

            internal readonly XLBookPoint Next;

            public Link(XLBookPoint previous, XLBookPoint next)
            {
                Previous = previous;
                Next = next;
            }
        }
    }
}
