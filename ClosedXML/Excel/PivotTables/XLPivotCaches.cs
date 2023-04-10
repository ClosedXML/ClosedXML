using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotCaches : IXLPivotCaches, IEnumerable<XLPivotCache>
    {
        private readonly List<XLPivotCache> _caches = new();

        IXLPivotCache IXLPivotCaches.Add(IXLRange  range) => Add(range);

        IEnumerator<IXLPivotCache> IEnumerable<IXLPivotCache>.GetEnumerator() => GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal XLPivotCache Add(IXLRange range)
        {
            var newPivotCache = range is IXLTable table
                ? new XLPivotCache(table)
                : new XLPivotCache(range);

            newPivotCache.Refresh();
            _caches.Add(newPivotCache);
            return newPivotCache;
        }

        internal XLPivotCache Add(XLPivotSourceReference pivotSourceReference)
        {
            var newPivotSource = pivotSourceReference.SourceType switch
            {
                XLPivotTableSourceType.Table => new XLPivotCache(pivotSourceReference.SourceTable!),
                XLPivotTableSourceType.Range => new XLPivotCache(pivotSourceReference.SourceRange),
                _ => throw new NotSupportedException("Unexpected source type.")
            };
            _caches.Add(newPivotSource);
            return newPivotSource;
        }

        public IEnumerator<XLPivotCache> GetEnumerator() => _caches.GetEnumerator();

        internal IEnumerable<XLPivotCache> GetAll(IXLRange range)
        {
            return _caches.Where(s => Equals(s.PivotSourceReference.SourceRange, range));
        }
    }
}
