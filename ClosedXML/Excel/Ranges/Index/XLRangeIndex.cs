using ClosedXML.Excel.Patterns;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.Ranges.Index
{
    /// <summary>
    /// Implementation of <see cref="IXLRangeIndex"/> internally using QuadTree.
    /// </summary>
    internal abstract class XLRangeIndex : IXLRangeIndex
    {
        #region Public Constructors

        public XLRangeIndex(IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
            _rangeList = new List<IXLAddressable>();
            (_worksheet as XLWorksheet).RegisterRangeIndex(this);
        }

        #endregion Public Constructors

        #region Public Methods

        public abstract bool MatchesType(XLRangeType rangeType);

        public bool Add(IXLAddressable range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            if (!range.RangeAddress.IsValid)
                throw new ArgumentException("Range is invalid");

            CheckWorksheet(range.RangeAddress.Worksheet);

            _count++;
            if (_count < MinimumCountForIndexing)
            {
                if (_rangeList.Any(r => r == range))
                    return false;

                _rangeList.Add(range);
                return true;
            }

            if (_quadTree == null)
                InitializeTree();

            return _quadTree.Add(range);
        }

        public bool Contains(in XLAddress address)
        {
            CheckWorksheet(address.Worksheet);

            if (_quadTree == null)
            {
                var addr = address;
                return _rangeList.Any(r => r.RangeAddress.Contains(addr));
            }

            return _quadTree.GetIntersectedRanges(address).Any();
        }

        public IEnumerable<IXLAddressable> GetAll()
        {
            if (_quadTree == null)
            {
                return _rangeList.AsEnumerable();
            }

            return _quadTree.GetAll();
        }

        public IEnumerable<IXLAddressable> GetIntersectedRanges(XLRangeAddress rangeAddress)
        {
            CheckWorksheet(rangeAddress.Worksheet);

            if (_quadTree == null)
            {
                return _rangeList.Where(r => r.RangeAddress.Intersects(rangeAddress));
            }

            return _quadTree.GetIntersectedRanges(rangeAddress);
        }

        public IEnumerable<IXLAddressable> GetIntersectedRanges(XLAddress address)
        {
            CheckWorksheet(address.Worksheet);

            if (_quadTree == null)
            {
                return _rangeList.Where(r => r.RangeAddress.Contains(address));
            }

            return _quadTree.GetIntersectedRanges(address);
        }

        public bool Intersects(in XLRangeAddress rangeAddress)
        {
            CheckWorksheet(rangeAddress.Worksheet);

            if (_quadTree == null)
            {
                var addr = rangeAddress;
                return _rangeList.Any(r => r.RangeAddress.Intersects(addr));
            }

            return _quadTree.GetIntersectedRanges(rangeAddress).Any();
        }

        public bool Remove(IXLRangeAddress rangeAddress)
        {
            if (rangeAddress == null)
                throw new ArgumentNullException(nameof(rangeAddress));

            CheckWorksheet(rangeAddress.Worksheet);

            if (_quadTree == null)
            {
                return _rangeList.RemoveAll(r => Equals(r.RangeAddress, rangeAddress)) > 0;
            }

            return _quadTree.Remove(rangeAddress);
        }

        public int RemoveAll(Predicate<IXLAddressable> predicate = null)
        {
            predicate = predicate ?? (_ => true);

            if (_quadTree == null)
            {
                return _rangeList.RemoveAll(predicate);
            }

            return _quadTree.RemoveAll(predicate).Count();
        }

        #endregion Public Methods

        #region Private Fields

        /// <summary>
        /// The minimum number of ranges to be included into a QuadTree. Until it is reached the ranges
        /// are added into a simple list to minimize the overhead of searching intersections on small collections.
        /// </summary>
        private const int MinimumCountForIndexing = 20;

        /// <summary>
        /// A collection of ranges used before the QuadTree is initialized (until <see cref="MinimumCountForIndexing"/>
        /// is reached.
        /// </summary>
        protected readonly List<IXLAddressable> _rangeList;

        private readonly IXLWorksheet _worksheet;
        private int _count = 0;
        protected Quadrant _quadTree;

        #endregion Private Fields

        #region Private Methods

        private void CheckWorksheet(IXLWorksheet worksheet)
        {
            if (worksheet != _worksheet)
                throw new ArgumentException("Range belongs to a different worksheet");
        }

        private void InitializeTree()
        {
            _quadTree = CreateQuadTree();
            _rangeList.ForEach(r => _quadTree.Add(r));
            _rangeList.Clear();
        }

        protected virtual Quadrant CreateQuadTree()
        {
            return new Quadrant();
        }

        #endregion Private Methods
    }

    /// <summary>
    /// Generic version of <see cref="XLRangeIndex"/>.
    /// </summary>
    internal class XLRangeIndex<T> : XLRangeIndex, IXLRangeIndex<T>
        where T : IXLAddressable
    {
        public XLRangeIndex(IXLWorksheet worksheet) : base(worksheet)
        {
        }

        public bool Add(T range)
        {
            return base.Add(range);
        }

        public int RemoveAll(Predicate<T> predicate)
        {
            predicate = predicate ?? (_ => true);

            return base.RemoveAll(r => predicate((T)r));
        }

        public new IEnumerable<T> GetIntersectedRanges(XLRangeAddress rangeAddress)
        {
            return base.GetIntersectedRanges(rangeAddress).Cast<T>();
        }

        public new IEnumerable<T> GetIntersectedRanges(XLAddress address)
        {
            return base.GetIntersectedRanges(address).Cast<T>();
        }

        public override bool MatchesType(XLRangeType rangeType)
        {
            var innerType = typeof(T);

            if (innerType == typeof(IXLRangeBase) ||
                innerType == typeof(XLRangeBase))
                return true;

            switch (rangeType)
            {
                case XLRangeType.Range:
                    return innerType == typeof(IXLRange) ||
                           innerType == typeof(XLRange);

                case XLRangeType.Column:
                    return innerType == typeof(IXLColumn) ||
                           innerType == typeof(XLColumn);

                case XLRangeType.Row:
                    return innerType == typeof(IXLRow) ||
                           innerType == typeof(XLRow);

                case XLRangeType.RangeColumn:
                    return innerType == typeof(IXLRangeColumn) ||
                           innerType == typeof(XLRangeColumn);

                case XLRangeType.RangeRow:
                    return innerType == typeof(IXLRangeRow) ||
                           innerType == typeof(XLRangeRow);

                case XLRangeType.Table:
                    return innerType == typeof(IXLTable) ||
                           innerType == typeof(XLTable);

                case XLRangeType.Worksheet:
                    return innerType == typeof(IXLWorksheet) ||
                           innerType == typeof(XLWorksheet);

                default:
                    throw new NotImplementedException(nameof(rangeType));
            }
        }

        public new IEnumerable<T> GetAll()
        {
            return base.GetAll().Cast<T>();
        }

        protected override Quadrant CreateQuadTree()
        {
            return new Quadrant<T>();
        }
    }
}
