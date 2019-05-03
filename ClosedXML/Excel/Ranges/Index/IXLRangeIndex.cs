using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Ranges.Index
{
    /// <summary>
    /// Interface for the engine aimed to speed-up the search for the range intersections.
    /// </summary>
    internal interface IXLRangeIndex
    {
        bool Add(IXLRangeBase range);

        bool Remove(IXLRangeAddress rangeAddress);

        int RemoveAll(Predicate<IXLRangeBase> predicate = null);

        IEnumerable<IXLRangeBase> GetIntersectedRanges(XLRangeAddress rangeAddress);

        IEnumerable<IXLRangeBase> GetIntersectedRanges(XLAddress address);

        IEnumerable<IXLRangeBase> GetAll();

        bool Intersects(in XLRangeAddress rangeAddress);

        bool Contains(in XLAddress address);

        bool MatchesType(XLRangeType rangeType);
    }

    internal interface IXLRangeIndex<T> : IXLRangeIndex
        where T : IXLRangeBase
    {
        bool Add(T range);

        int RemoveAll(Predicate<T> predicate = null);

        new IEnumerable<T> GetIntersectedRanges(XLRangeAddress rangeAddress);

        new IEnumerable<T> GetIntersectedRanges(XLAddress address);

        new IEnumerable<T> GetAll();
    }
}
