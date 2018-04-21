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
            _quadTree = new Quadrant();
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

            return _quadTree.Add(range);
        }

        public bool Contains(in XLAddress address)
        {
            CheckWorksheet(address.Worksheet);
            return _quadTree.GetIntersectedRanges(address).Any();
        }

        public bool Intersects(in XLRangeAddress rangeAddress)
        {
            CheckWorksheet(rangeAddress.Worksheet);
            return _quadTree.GetIntersectedRanges(rangeAddress).Any();
        }

        public bool Remove(IXLRangeBase range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            CheckWorksheet(range.Worksheet);

            return _quadTree.Remove(range);
        }

        public int RemoveAll(Predicate<IXLRangeBase> predicate = null)
        {
            return _quadTree.RemoveAll(predicate ?? (_ => true)).Count();
        }

        public IEnumerable<IXLRangeBase> GetAll()
        {
            return _quadTree.GetAll();
        }

        public IEnumerable<IXLRangeBase> GetIntersectedRanges(XLRangeAddress rangeAddress)
        {
            CheckWorksheet(rangeAddress.Worksheet);

            return _quadTree.GetIntersectedRanges(rangeAddress);
        }

        public IEnumerable<IXLRangeBase> GetIntersectedRanges(XLAddress address)
        {
            CheckWorksheet(address.Worksheet);

            return _quadTree.GetIntersectedRanges(address);
        }

        #endregion Public Methods

        #region Private Fields

        private readonly Quadrant _quadTree;
        private readonly IXLWorksheet _worksheet;

        #endregion Private Fields

        #region Private Methods

        private void CheckWorksheet(IXLWorksheet worksheet)
        {
            if (worksheet != _worksheet)
                throw new ArgumentException("Range belongs to a different worksheet");
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
