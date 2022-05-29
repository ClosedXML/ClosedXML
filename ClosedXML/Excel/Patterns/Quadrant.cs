using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.Patterns
{
    /// <summary>
    /// Implementation of QuadTree adapted to Excel worksheet specifics. Differences with the classic implementation
    /// are that the topmost level is split to 128 square parts (2 columns of 64 blocks, each 8192*8192 cells) and that splitting
    /// the quadrant onto 4 smaller quadrants does not depend on the number of items in this quadrant. When the range is added to the
    /// QuadTree it is placed on the bottommost level where it fits to a single quadrant. That means, row-wide and column-wide ranges
    /// are always placed at the level 0, and the smaller the range is the deeper it goes down the tree. This approach eliminates
    /// the need of transferring ranges between levels.
    /// </summary>
    internal class Quadrant
    {
        #region Public Properties

        /// <summary>
        /// Smaller quadrants which the current one is split to. Is NULL until ranges are added to child quadrants.
        /// </summary>
        public IEnumerable<Quadrant> Children { get; private set; }

        /// <summary>
        /// The level of current quadrant. Top most has level 0, child quadrants has levels (Level + 1).
        /// </summary>
        public byte Level { get; }

        /// <summary>
        /// Minimum column included in this quadrant.
        /// </summary>
        public int MinimumColumn { get; }

        /// <summary>
        /// Minimum row included in this quadrant.
        /// </summary>
        public int MinimumRow { get; }

        /// <summary>
        /// Maximum column included in this quadrant.
        /// </summary>
        public int MaximumColumn { get; }

        /// <summary>
        /// Maximum row included in this quadrant.
        /// </summary>
        public int MaximumRow { get; }

        /// <summary>
        /// Collection of ranges belonging to this quadrant (does not include ranges from child quadrants).
        /// </summary>
        public IEnumerable<IXLAddressable> Ranges
        {
            get => _ranges?.Values.AsEnumerable();
        }

        /// <summary>
        /// The number of current quadrant by horizontal axis.
        /// </summary>
        public short X { get; private set; }

        /// <summary>
        /// The number of current quadrant by vertical axis.
        /// </summary>
        public short Y { get; private set; }

        #endregion Public Properties

        #region Constructors

        public Quadrant() : this(0, 0, 0)
        { }

        private Quadrant(byte level, short x, short y)
        {
            Level = level;
            X = x;
            Y = y;

            MinimumColumn = (Level == 0) ? 1 : 1 + XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * X;
            MinimumRow = (Level == 0) ? 1 : 1 + XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * Y; //MaxColumnNumber here is not a mistake
            MaximumColumn = (Level == 0)
                ? XLHelper.MaxColumnNumber
                : XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * (X + 1);
            MaximumRow = (Level == 0)
                ? XLHelper.MaxRowNumber
                : XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * (Y + 1); //MaxColumnNumber here is not a mistake
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Add a range to the quadrant or to one of the child quadrants (recursively).
        /// </summary>
        /// <returns>True, if range was successfully added, false if it has been added before.</returns>
        public bool Add(IXLAddressable range)
        {
            bool res = false;
            var children = Children ?? CreateChildren().ToList();
            bool addToChild = false;
            foreach (var childQuadrant in children)
            {
                var rangeAddress = range.RangeAddress;
                if (childQuadrant.Covers(in rangeAddress))
                {
                    res |= childQuadrant.Add(range);
                    addToChild = true;
                    break;
                }
            }

            if (!addToChild)
                res = AddInternal(range);

            if (Children == null && addToChild)
                Children = children;

            return res;
        }

        /// <summary>
        /// Get all ranges from the quadrant and all child quadrants (recursively).
        /// </summary>
        public IEnumerable<IXLAddressable> GetAll()
        {
            if (Ranges != null)
            {
                foreach (var range in Ranges)
                    yield return range;
            }

            if (Children != null)
            {
                foreach (var childQuadrant in Children)
                {
                    var childRanges = childQuadrant.GetAll();
                    foreach (var range in childRanges)
                        yield return range;
                }
            }
        }

        /// <summary>
        /// Get all ranges from the quadrant and all child quadrants (recursively) that intersect the specified address.
        /// </summary>
        public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLRangeAddress rangeAddress)
        {
            if (Ranges != null)
            {
                foreach (var range in Ranges)
                {
                    if (range.RangeAddress.Intersects(rangeAddress))
                        yield return range;
                }
            }

            if (Children != null)
            {
                foreach (var childQuadrant in Children)
                {
                    if (childQuadrant.Intersects(in rangeAddress))
                    {
                        var childRanges = childQuadrant.GetIntersectedRanges(rangeAddress);
                        foreach (var range in childRanges)
                            yield return range;
                    }
                }
            }
        }

        /// <summary>
        /// Get all ranges from the quadrant and all child quadrants (recursively) that cover the specified address.
        /// </summary>
        public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLAddress address)
        {
            if (Ranges != null)
            {
                foreach (var range in Ranges)
                {
                    if (range.RangeAddress.Contains(address))
                        yield return range;
                }
            }

            if (Children != null)
            {
                foreach (var childQuadrant in Children)
                {
                    if (childQuadrant.Covers(in address))
                    {
                        var childRanges = childQuadrant.GetIntersectedRanges(address);
                        foreach (var range in childRanges)
                            yield return range;
                    }
                }
            }
        }

        /// <summary>
        /// Remove the range from the quadrant or from child quadrants (recursively).
        /// </summary>
        /// <returns>True if the range was removed, false if it does not exist in the QuadTree.</returns>
        public bool Remove(IXLRangeAddress rangeAddress)
        {
            bool res = false;

            bool coveredByChild = false;
            if (Children != null)
            {
                foreach (var childQuadrant in Children)
                {
                    if (childQuadrant.Covers(rangeAddress))
                    {
                        res |= childQuadrant.Remove(rangeAddress);
                        coveredByChild = true;
                    }
                }
            }

            if (!coveredByChild)
            {
                if (_ranges?.Remove(rangeAddress) == true)
                    res = true;
            }

            return res;
        }

        /// <summary>
        /// Remove all the ranges matching specified criteria from the quadrant and its child quadrants (recursively).
        /// Don't use it for searching intersections as it would be much less efficient than <see cref="GetIntersectedRanges(IXLRangeAddress)"/>.
        /// </summary>
        public IEnumerable<IXLAddressable> RemoveAll(Predicate<IXLAddressable> predicate)
        {
            if (_ranges != null)
            {
                var ranges = _ranges.Values.Where(r => predicate(r));
                var keysToRemove = new List<IXLRangeAddress>();
                foreach (var range in ranges)
                {
                    keysToRemove.Add(range.RangeAddress);
                    yield return range;
                }

                foreach (var keyToRemove in keysToRemove)
                {
                    _ranges.Remove(keyToRemove);
                }
            }

            if (Children != null)
            {
                foreach (var childQuadrant in Children)
                    foreach (var childRange in childQuadrant.RemoveAll(predicate))
                    {
                        yield return childRange;
                    }
            }
        }

        #endregion Public Methods

        #region Private Fields

        /// <summary>
        /// Maximum depth of the QuadTree. Value 10 corresponds to the smallest quadrants having size 16*16 cells.
        /// </summary>
        private const byte MAX_LEVEL = 10;

        /// <summary>
        /// Collection of ranges belonging to the current quadrant (that cannot fit into child quadrants).
        /// </summary>
        private Dictionary<IXLRangeAddress, IXLAddressable> _ranges;

        #endregion Private Fields

        #region Private Methods

        /// <summary>
        /// Add a range to the collection of quadrant's own ranges.
        /// </summary>
        /// <returns>True if the range was successfully added, false if it had been added before.</returns>
        private bool AddInternal(IXLAddressable range)
        {
            if (_ranges == null)
                _ranges = new Dictionary<IXLRangeAddress, IXLAddressable>();

            if (_ranges.ContainsKey(range.RangeAddress))
                return false;

            _ranges.Add(range.RangeAddress, range);
            return true;
        }

        /// <summary>
        /// Check if the current quadrant fully covers the specified address.
        /// </summary>
        private bool Covers(in IXLRangeAddress rangeAddress)
        {
            return MinimumColumn <= rangeAddress.FirstAddress.ColumnNumber &&
                   MaximumColumn >= rangeAddress.LastAddress.ColumnNumber &&
                   MinimumRow <= rangeAddress.FirstAddress.RowNumber &&
                   MaximumRow >= rangeAddress.LastAddress.RowNumber;
        }

        /// <summary>
        /// Check if the current quadrant covers the specified address.
        /// </summary>
        private bool Covers(in IXLAddress address)
        {
            return MinimumColumn <= address.ColumnNumber &&
                   MaximumColumn >= address.ColumnNumber &&
                   MinimumRow <= address.RowNumber &&
                   MaximumRow >= address.RowNumber;
        }

        /// <summary>
        /// Check if the current quadrant intersects the specified address.
        /// </summary>
        private bool Intersects(in IXLRangeAddress rangeAddress)
        {
            return ((MinimumRow <= rangeAddress.FirstAddress.RowNumber && rangeAddress.FirstAddress.RowNumber <= MaximumRow) ||
                    (rangeAddress.FirstAddress.RowNumber <= MinimumRow && MinimumRow <= rangeAddress.LastAddress.RowNumber))
                   &&
                   ((MinimumColumn <= rangeAddress.FirstAddress.ColumnNumber && rangeAddress.FirstAddress.ColumnNumber <= MaximumColumn) ||
                    (rangeAddress.FirstAddress.ColumnNumber <= MinimumColumn && MinimumColumn <= rangeAddress.LastAddress.ColumnNumber));
        }

        /// <summary>
        /// Create a collection of child quadrants dividing the current one.
        /// </summary>
        private IEnumerable<Quadrant> CreateChildren()
        {
            byte childLevel = (byte)(Level + 1);
            if (childLevel > MAX_LEVEL)
                yield break;
            byte xCount = 2; // Always divide on halves
            byte yCount = (byte)((Level == 0) ? (XLHelper.MaxRowNumber / XLHelper.MaxColumnNumber) : 2); // Level 0 divide onto 64 parts, the rest - on halves

            for (byte dy = 0; dy < yCount; dy++)
            {
                for (byte dx = 0; dx < xCount; dx++)
                {
                    yield return new Quadrant(childLevel, (short)(X * 2 + dx), (short)(Y * 2 + dy));
                }
            }
        }

        #endregion Private Methods
    }

    /// <summary>
    /// A generic version of <see cref="Quadrant"/>
    /// </summary>
    internal class Quadrant<T> : Quadrant
        where T:IXLAddressable
    {
        public new IEnumerable<T> Ranges => base.Ranges.Cast<T>();

        public bool Add(T range)
        {
            return base.Add(range);
        }

        public new IEnumerable<T> GetAll()
        {
            return base.GetAll().Cast<T>();
        }

        public new IEnumerable<T> GetIntersectedRanges(IXLRangeAddress rangeAddress)
        {
            return base.GetIntersectedRanges(rangeAddress).Cast<T>();
        }

        public new IEnumerable<T> GetIntersectedRanges(IXLAddress address)
        {
            return base.GetIntersectedRanges(address).Cast<T>();
        }

        public bool Remove(T range)
        {
            return Remove(range.RangeAddress);
        }
        public IEnumerable<T> RemoveAll(Predicate<T> predicate)
        {
            return base.RemoveAll(r => predicate((T) r)).Cast<T>();
        }
    }
}
