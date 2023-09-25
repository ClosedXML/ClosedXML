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
    /// <see cref="MoveToCurrent"/> and <see cref="Reset"/>, but only one traversal
    /// can go on at the same time due to shared info about cycle detection.
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
        /// 1 based position of <see cref="_current"/>, if there is a traversal
        /// in progress (0 otherwise).
        /// </summary>
        private int _currentPosition;

        /// <summary>
        /// The address of a current of the chain.
        /// </summary>
        internal XLBookPoint Current => _current!.Value;

        /// <summary>
        /// Is there a cycle in the chain? Detected when a link has appeared
        /// as a current more than once and the current hasn't moved in the
        /// meantime.
        /// </summary>
        internal bool IsCurrentInCycle { get; private set; }

        /// <summary>
        /// Create a new chain filled with all formulas from the workbook.
        /// </summary>
        internal static XLCalculationChain CreateFrom(XLWorkbook wb)
        {
            var chain = new XLCalculationChain();
            foreach (var sheet in wb.WorksheetsInternal)
            {
                var formulaSlice = sheet.Internals.CellsCollection.FormulaSlice;
                using var e = formulaSlice.GetForwardEnumerator(XLSheetRange.Full);
                while (e.MoveNext())
                    chain.AddLast(new XLBookPoint(sheet.SheetId, e.Point));
            }

            return chain;
        }

        /// <summary>
        /// Add a new link at the beginning of a chain.
        /// </summary>
        private void AddFirst(XLBookPoint point, int lastPosition)
        {
            if (_head is null || _tail is null)
            {
                Init(point);
                return;
            }

            Insert(point, lastPosition, _tail.Value, _head.Value);
            _head = point;
        }

        /// <inheritdoc cref="AddLast(XLBookPoint,int)"/>
        internal void AddLast(XLBookPoint point) => AddLast(point, 0);

        /// <summary>
        /// Add all cells from the area to the end of the chain.
        /// </summary>
        /// <exception cref="ArgumentException">If chain already contains a cell from the area.</exception>
        internal void AppendArea(uint sheetId, XLSheetRange range)
        {
            for (var row = range.TopRow; row <= range.BottomRow; ++row)
            {
                for (var col = range.LeftColumn; col <= range.RightColumn; ++col)
                {
                    AddLast(new XLBookPoint(sheetId, new XLSheetPoint(row, col)));
                }
            }
        }

        /// <summary>
        /// Append formula at the end of the chain.
        /// </summary>
        private void AddLast(XLBookPoint point, int lastPosition)
        {
            if (_head is null || _tail is null)
            {
                Init(point);
                return;
            }

            Insert(point, lastPosition, _tail.Value, _head.Value);
            _tail = point;
        }

        /// <summary>
        /// Initialize empty chain with a single link chain.
        /// </summary>
        private void Init(XLBookPoint point)
        {
            Debug.Assert(_nodeMap.Count == 0 && _head is null && _tail is null);
            _nodeMap.Add(point, new Link(point, point, 0));
            _head = _tail = point;
        }

        /// <summary>
        /// Insert a link into the <see cref="_nodeMap"/> between
        /// <paramref name="prev"/> and <paramref name="next"/>.
        /// Don't update head or tail.
        /// </summary>
        private void Insert(XLBookPoint point, int lastPosition, XLBookPoint prev, XLBookPoint next)
        {
            _nodeMap.Add(point, new Link(prev, next, lastPosition));

            var prevLink = _nodeMap[prev];
            _nodeMap[prev] = new Link(prevLink.Previous, point, prevLink.LastPosition);

            var nextLink = _nodeMap[next];
            _nodeMap[next] = new Link(point, nextLink.Next, nextLink.LastPosition);
        }

        /// <summary>
        /// Add a link for <paramref name="point"/> after the link for
        /// <paramref name="anchor"/>.
        /// </summary>
        /// <param name="anchor">
        /// The anchor point after which will be the new point added.
        /// </param>
        /// <param name="point">Point to add to the chain.</param>
        /// <param name="lastPosition">The last position of the point in the chain.</param>
        internal void AddAfter(XLBookPoint anchor, XLBookPoint point, int lastPosition)
        {
            var prevLink = _nodeMap[anchor];
            var next = prevLink.Next;
            Insert(point, lastPosition, anchor, next);

            if (anchor == _tail!.Value)
                _tail = point;
        }

        /// <summary>
        /// Remove point from the chain.
        /// </summary>
        /// <param name="point">Link to remove.</param>
        /// <returns>Last position of the removed link.</returns>
        /// <exception cref="InvalidOperationException">Point is not a part of the chain.</exception>
        internal int Remove(XLBookPoint point)
        {
            if (!_nodeMap.TryGetValue(point, out var pointLink))
                throw PointNotInChain(point);

            // Point is in the chain and there is exactly one link -> clear all.
            if (_nodeMap.Count == 1)
            {
                Clear();
                return pointLink.LastPosition;
            }

            if (point == _head!.Value)
                _head = pointLink.Next;

            if (point == _tail!.Value)
                _tail = pointLink.Previous;

            var prevLink = _nodeMap[pointLink.Previous];
            Debug.Assert(prevLink.Next == point);
            _nodeMap[pointLink.Previous] = new Link(prevLink.Previous, pointLink.Next, prevLink.LastPosition);

            var nextLink = _nodeMap[pointLink.Next];
            Debug.Assert(nextLink.Previous == point);
            _nodeMap[pointLink.Next] = new Link(pointLink.Previous, nextLink.Next, nextLink.LastPosition);

            _nodeMap.Remove(point);
            return pointLink.LastPosition;
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
        /// Enumerate all links in the chain.
        /// </summary>
        internal IEnumerable<(XLBookPoint Point, int LastPosition)> GetLinks()
        {
            if (_head is null)
                yield break;

            var current = _head.Value;
            do
            {
                var link = _nodeMap[current];
                yield return new ValueTuple<XLBookPoint, int>(current, link.LastPosition);
                current = link.Next;
            } while (current != _head.Value);
        }

        internal void Reset()
        {
            if (_current is null)
                return;

            var point = _current.Value;
            var link = _nodeMap[point];
            while (link.LastPosition != 0)
            {
                _nodeMap[point] = new Link(link.Previous, link.Next, 0);
                point = link.Next;
                link = _nodeMap[point];
            }

            _current = null;
            _currentPosition = 0;
        }

        /// <summary>
        /// Mark current link as complete and move ahead to the next link.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the enumerator moved ahead, <c>false</c> if
        /// there are no more links and chain has looped completely.
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
                _currentPosition = 1;
                return true;
            }

            // Subsequent move
            var currentPoint = _current.Value;
            if (!_nodeMap.TryGetValue(currentPoint, out var currentLink))
                throw PointNotInChain(currentPoint);

            // Clear up the last position, the current point is being moved to done
            // and clearing will ensure next traversal won't be affected.
            if (currentLink.LastPosition != 0)
                _nodeMap[currentPoint] = new Link(currentLink.Previous, currentLink.Next, 0);

            var nextPoint = currentLink.Next;
            Debug.Assert(_nodeMap[nextPoint].Previous == currentPoint);
            if (nextPoint == _head!.Value)
            {
                // Whole chain has been calculated.
                return false;
            }

            // Since we moved, the new last position is greater than all others
            // and thus can't be in the cycle.
            IsCurrentInCycle = false;
            _current = nextPoint;
            _currentPosition++;
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
            {
                // But it basically means that currentPoint depends on pointToMove
                // thus cell depends on itself and that is a cycle.
                IsCurrentInCycle = true;
                return;
            }

            // If head is also current, moving before the current means moving before head
            var pointToMoveLastPosition = Remove(pointToMove);
            if (_head == currentPoint)
            {
                AddFirst(pointToMove, pointToMoveLastPosition);
            }
            else
            {
                // Current is not a head = move a link after prev of current.
                var anchor = _nodeMap[currentPoint].Previous;
                AddAfter(anchor, pointToMove, pointToMoveLastPosition);
            }

            var shiftedLink = _nodeMap[currentPoint];
            _nodeMap[currentPoint] = new Link(shiftedLink.Previous, shiftedLink.Next, _currentPosition);

            IsCurrentInCycle = _currentPosition == pointToMoveLastPosition;
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

            /// <summary>
            /// <para>
            /// What was the 1-based position of the link in the chain the last
            /// time the link has been current. Only used when link is pushed
            /// to the back, otherwise it's <c>0</c>.
            /// </para>
            /// <para>
            /// The last position of a link is only updated when
            /// <list type="bullet">
            /// <item>
            /// Link is moved from current to the back - that means link
            /// will be moved to current again at some point in the future
            /// and if chain hasn't processed even one link in the meantime,
            /// there is a cycle.
            /// </item>
            /// <item>
            /// Link is marked as done and current moves past it. The last
            /// position should be cleared as not to confuse next traversal.
            /// </item>
            /// <item>
            /// Chain traversal is reset - links in front of current may still
            /// have set their last position, because other links have been
            /// moved to the current as a supporting links.
            /// </item>
            /// </list>
            /// </para>
            /// </summary>
            /// <remarks>Used for cycle detection.</remarks>
            internal readonly int LastPosition;

            public Link(XLBookPoint previous, XLBookPoint next, int lastPosition)
            {
                Previous = previous;
                Next = next;
                LastPosition = lastPosition;
            }
        }
    }
}
