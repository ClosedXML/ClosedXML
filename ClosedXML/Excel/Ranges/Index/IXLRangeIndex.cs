using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Ranges.Index
{
    /// <summary>
    /// Interface for the engine aimed to speed-up the search for the range intersections.
    /// </summary>
    internal interface IXLRangeIndex
    {
        bool Add(IXLAddressable range);

        bool Remove(IXLRangeAddress rangeAddress);

        int RemoveAll(Predicate<IXLAddressable> predicate = null);

        IEnumerable<IXLAddressable> GetIntersectedRanges(XLRangeAddress rangeAddress);

        IEnumerable<IXLAddressable> GetIntersectedRanges(XLAddress address);

        IEnumerable<IXLAddressable> GetAll();

        bool Intersects(in XLRangeAddress rangeAddress);

        bool Contains(in XLAddress address);

        bool MatchesType(XLRangeType rangeType);
    }

    internal interface IXLRangeIndex<T> : IXLRangeIndex
        where T : IXLAddressable
    {
        bool Add(T range);

        int RemoveAll(Predicate<T> predicate = null);

        new IEnumerable<T> GetIntersectedRanges(XLRangeAddress rangeAddress);

        new IEnumerable<T> GetIntersectedRanges(XLAddress address);

        new IEnumerable<T> GetAll();
    }
}
