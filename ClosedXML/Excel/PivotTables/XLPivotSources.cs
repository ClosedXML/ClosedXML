using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    public interface IXLPivotSources : IEnumerable<IXLPivotSource>
    {
        IXLPivotSource Add(IXLPivotSourceReference pivotSourceReference);

        IXLPivotSource Add(IXLRange range);

        IXLPivotSource Add(IXLTable table);

        Boolean Contains(IXLRange range);

        Boolean Contains(IXLTable table);

        Boolean TryGet(IXLRange range, out IXLPivotSource pivotSource);

        Boolean TryGet(IXLTable table, out IXLPivotSource pivotSource);
    }

    internal class XLPivotSources : IXLPivotSources
    {
        private readonly IDictionary<IXLRange, IXLPivotSource> _rangePivotSources = new Dictionary<IXLRange, IXLPivotSource>();
        private readonly IDictionary<String, IXLPivotSource> _tablePivotSources = new Dictionary<String, IXLPivotSource>(StringComparer.OrdinalIgnoreCase);

        public IXLPivotSource Add(IXLRange range)
        {
            return Add(range, refresh: true);
        }

        public IXLPivotSource Add(IXLRange range, bool refresh)
        {
            if (_rangePivotSources.TryGetValue(range, out IXLPivotSource pivotSource))
                return pivotSource;
            else
            {
                var newPivotSource = new XLPivotSource(range);
                if (refresh)
                    newPivotSource.Refresh();
                _rangePivotSources.Add(range, newPivotSource);
                return newPivotSource;
            }
        }

        public IXLPivotSource Add(IXLTable table)
        {
            return Add(table, refresh: true);
        }

        public IXLPivotSource Add(IXLTable table, bool refresh)
        {
            if (_tablePivotSources.TryGetValue(table.Name, out IXLPivotSource pivotSource))
                return pivotSource;
            else
            {
                var newPivotSource = new XLPivotSource(table);
                if (refresh)
                    newPivotSource.Refresh();
                _tablePivotSources.Add(table.Name, newPivotSource);
                return newPivotSource;
            }
        }

        public IXLPivotSource Add(IXLPivotSourceReference pivotSourceReference)
        {
            return Add(pivotSourceReference, refresh: true);
        }

        public IXLPivotSource Add(IXLPivotSourceReference pivotSourceReference, bool refresh)
        {
            switch (pivotSourceReference.SourceType)
            {
                case XLPivotTableSourceType.Table:
                    return Add(pivotSourceReference.SourceTable, refresh);

                case XLPivotTableSourceType.Range:
                    return Add(pivotSourceReference.SourceRange, refresh);

                default:
                    throw new NotImplementedException();
            }
        }

        public bool Contains(IXLRange range)
        {
            return _rangePivotSources.ContainsKey(range);
        }

        public bool Contains(IXLTable table)
        {
            return _tablePivotSources.ContainsKey(table.Name);
        }

        public IEnumerator<IXLPivotSource> GetEnumerator()
        {
            return _rangePivotSources.Values
                .Union(_tablePivotSources.Values)
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool TryGet(IXLRange range, out IXLPivotSource pivotSource)
        {
            return _rangePivotSources.TryGetValue(range, out pivotSource);
        }

        public bool TryGet(IXLTable table, out IXLPivotSource pivotSource)
        {
            return _tablePivotSources.TryGetValue(table.Name, out pivotSource);
        }
    }
}
