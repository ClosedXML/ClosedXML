using ClosedXML.Excel.Patterns;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.Ranges.Index
{
    /// <summary>
    /// Implementation of <see cref="IXLRangeIndex"/> internally using QuadTree.
    /// </summary>
    internal class XLRangeIndex : IXLRangeIndex
    {
        #region Public Constructors

        public XLRangeIndex(IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
            _rangeList = new List<IXLRangeBase>();
        }

        #endregion Public Constructors

        #region Public Methods

        public bool Add(IXLRangeBase range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            if (!range.RangeAddress.IsValid)
                throw new ArgumentException("Range is invalid");

            CheckWorksheet(range.Worksheet);

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

        public IEnumerable<IXLRangeBase> GetAll()
        {
            if (_quadTree == null)
            {
                return _rangeList.AsEnumerable();
            }

            return _quadTree.GetAll();
        }

        public IEnumerable<IXLRangeBase> GetIntersectedRanges(XLRangeAddress rangeAddress)
        {
            CheckWorksheet(rangeAddress.Worksheet);

            if (_quadTree == null)
            {
                return _rangeList.Where(r => r.RangeAddress.Intersects(rangeAddress));
            }

            return _quadTree.GetIntersectedRanges(rangeAddress);
        }

        public IEnumerable<IXLRangeBase> GetIntersectedRanges(XLAddress address)
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

        public bool Remove(IXLRangeBase range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            CheckWorksheet(range.Worksheet);

            if (_quadTree == null)
            {
                return _rangeList.Remove(range);
            }

            return _quadTree.Remove(range);
        }

        public int RemoveAll(Predicate<IXLRangeBase> predicate = null)
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
        private readonly List<IXLRangeBase> _rangeList;

        private readonly IXLWorksheet _worksheet;
        private int _count = 0;
        private Quadrant _quadTree;

        #endregion Private Fields

        #region Private Methods

        private void CheckWorksheet(IXLWorksheet worksheet)
        {
            if (worksheet != _worksheet)
                throw new ArgumentException("Range belongs to a different worksheet");
        }

        private void InitializeTree()
        {
            _quadTree = new Quadrant();
            _rangeList.ForEach(r => _quadTree.Add(r));
            _rangeList.Clear();
        }

        #endregion Private Methods
    }

    /// <summary>
    /// Generic version of <see cref="XLRangeIndex"/>.
    /// </summary>
    internal class XLRangeIndex<T> : XLRangeIndex, IXLRangeIndex<T>
        where T : IXLRangeBase
    {
        public XLRangeIndex(IXLWorksheet worksheet) : base(worksheet)
        {
        }

        public bool Add(T range)
        {
            return base.Add(range);
        }

        public bool Remove(T range)
        {
            return base.Remove(range);
        }

        public int RemoveAll(Predicate<T> predicate)
        {
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

        public new IEnumerable<T> GetAll()
        {
            return base.GetAll().Cast<T>();
        }
    }
}
