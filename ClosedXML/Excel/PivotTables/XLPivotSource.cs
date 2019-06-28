using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotSource : IXLPivotSource
    {
        public XLPivotSource(IXLRange sourceRange)
            : this()
        {
            this.PivotSourceReference = new XLPivotSourceReference { SourceRange = sourceRange };
        }

        public XLPivotSource(IXLTable table)
            : this()
        {
            this.PivotSourceReference = new XLPivotSourceReference { SourceTable = table };
        }

        private XLPivotSource()
        {
            this.Guid = Guid.NewGuid();
            CachedFields = new Dictionary<String, IList<Object>>(StringComparer.OrdinalIgnoreCase);
            SetExcelDefaults();
        }

        public IDictionary<String, IList<Object>> CachedFields { get; internal set; }
        public Guid Guid { get; private set; }
        public XLItemsToRetain ItemsToRetainPerField { get; set; }

        public IXLPivotSourceReference PivotSourceReference { get; set; }
        public Boolean RefreshDataOnOpen { get; set; }

        public Boolean SaveSourceData { get; set; }

        public IList<String> SourceRangeFields
        {
            get
            {
                return this.PivotSourceReference.SourceRange
                  .FirstRow()
                  .Cells()
                  .Select(c => c.GetString())
                  .ToList()
                  .AsReadOnly();
            }
        }

        internal uint? CacheId { get; set; }
        internal String WorkbookCacheRelId { get; set; }

        public IXLPivotSource Refresh()
        {
            CachedFields.Clear();

            foreach (var column in PivotSourceReference.SourceRange.Columns())
            {
                var header = column.FirstCell().GetFormattedString();
                var firstCellAddress = column.FirstCell().Address;
                var values = column.CellsUsed(c => !c.Address.Equals(firstCellAddress))
                    .Select(c => c.Value)
                    .Distinct()
                    .ToList();

                CachedFields.Add(AdjustedFieldName(header), values);
            }

            return this;
        }

        public IXLPivotSource SetItemsToRetainPerField(XLItemsToRetain value) { ItemsToRetainPerField = value; return this; }

        public IXLPivotSource SetRefreshDataOnOpen() { RefreshDataOnOpen = true; return this; }

        public IXLPivotSource SetRefreshDataOnOpen(Boolean value) { RefreshDataOnOpen = value; return this; }

        public IXLPivotSource SetSaveSourceData() { SaveSourceData = true; return this; }

        public IXLPivotSource SetSaveSourceData(Boolean value) { SaveSourceData = value; return this; }

        internal XLPivotSource AddCachedField(String fieldName, List<Object> items)
        {
            var cachedFields = CachedFields as Dictionary<String, IList<Object>>;
            cachedFields.Add(fieldName, items);
            return this;
        }

        private string AdjustedFieldName(string header)
        {
            var modifiedHeader = header;
            int i = 1;
            while (CachedFields.ContainsKey(modifiedHeader))
            {
                i++;
                modifiedHeader = header + i.ToInvariantString();
            }

            return modifiedHeader;
        }

        private void SetExcelDefaults()
        {
            SaveSourceData = true;
        }
    }
}
