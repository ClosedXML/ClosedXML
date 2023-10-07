using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPivotCaches : IXLPivotCaches, IEnumerable<XLPivotCache>
    {
        private readonly XLWorkbook _workbook;
        private readonly List<XLPivotCache> _caches = new();

        public XLPivotCaches(XLWorkbook workbook)
        {
            _workbook = workbook;
        }

        IXLPivotCache IXLPivotCaches.Add(IXLRange range) => Add(XLBookArea.From(range));

        IEnumerator<IXLPivotCache> IEnumerable<IXLPivotCache>.GetEnumerator() => GetEnumerator();

        IEnumerator<XLPivotCache> IEnumerable<XLPivotCache>.GetEnumerator() => GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public List<XLPivotCache>.Enumerator GetEnumerator() => _caches.GetEnumerator();

        internal XLPivotCache Add(XLBookArea area)
        {
            var sourceReference = _workbook.TryGetTable(area, out var table)
                ? new XLPivotSourceReference(table.Name)
                : new XLPivotSourceReference(area);

            var newPivotCache = new XLPivotCache(sourceReference, _workbook);
            newPivotCache.Refresh();
            _caches.Add(newPivotCache);
            return newPivotCache;
        }

        internal XLPivotCache Add(XLPivotSourceReference pivotSourceReference)
        {
            var newPivotCache = new XLPivotCache(pivotSourceReference, _workbook);
            _caches.Add(newPivotCache);
            return newPivotCache;
        }

        /// <summary>
        /// Try to find an existing pivot cache for the passed area. The area
        /// is checked against both types of source references (tables and
        /// ranges) and if area matches, the cache is returned.
        /// </summary>
        internal XLPivotCache? Find(XLBookArea area)
        {
            // This method mimics behavior of Excel.
            // If there is a table for the area and there is a cache for the table, return cache for the table.
            if (_workbook.TryGetTable(area, out var table))
            {
                // Table exists, so try to find it and match with the source reference.
                foreach (var cache in _caches)
                {
                    if (XLHelper.NameComparer.Equals(cache.PivotSourceReference.Name, table.Name))
                        return cache;
                }
            }

            // Try to find a cache with area source.
            foreach (var cache in _caches)
            {
                var cacheArea = cache.PivotSourceReference.Area;
                if (cacheArea is not null && cacheArea.Value == area)
                    return cache;
            }

            return null;
        }
    }
}
