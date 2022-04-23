using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        #region Nested type: SaveContext

        internal sealed class SaveContext
        {
            public SaveContext()
            {
                DifferentialFormats = new Dictionary<XLStyleValue, int>();
                PivotTables = new Dictionary<Guid, PivotTableInfo>();
                RelIdGenerator = new RelIdGenerator();
                SharedFonts = new Dictionary<XLFontValue, FontInfo>();
                SharedNumberFormats = new Dictionary<int, NumberFormatInfo>();
                SharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
                TableId = 0;
                TableNames = new HashSet<string>();
            }

            public Dictionary<XLStyleValue, int> DifferentialFormats { get; private set; }
            public IDictionary<Guid, PivotTableInfo> PivotTables { get; private set; }
            public RelIdGenerator RelIdGenerator { get; private set; }
            public Dictionary<XLFontValue, FontInfo> SharedFonts { get; private set; }
            public Dictionary<int, NumberFormatInfo> SharedNumberFormats { get; private set; }
            public Dictionary<XLStyleValue, StyleInfo> SharedStyles { get; private set; }
            public uint TableId { get; set; }
            public HashSet<string> TableNames { get; private set; }
        }

        #endregion Nested type: SaveContext

        #region Nested type: RelType

        internal enum RelType
        {
            Workbook//, Worksheet
        }

        #endregion Nested type: RelType

        #region Nested type: RelIdGenerator

        internal sealed class RelIdGenerator
        {
            private readonly Dictionary<RelType, List<string>> _relIds = new Dictionary<RelType, List<string>>();

            public void AddValues(IEnumerable<string> values, RelType relType)
            {
                if (!_relIds.TryGetValue(relType, out var list))
                {
                    list = new List<string>();
                    _relIds.Add(relType, list);
                }
                list.AddRange(values.Where(v => !list.Contains(v)));
            }

            public string GetNext()
            {
                return GetNext(RelType.Workbook);
            }

            public string GetNext(RelType relType)
            {
                if (!_relIds.TryGetValue(relType, out var list))
                {
                    list = new List<string>();
                    _relIds.Add(relType, list);
                }

                var id = list.Count + 1;
                while (true)
                {
                    var relId = string.Concat("rId", id);
                    if (!list.Contains(relId))
                    {
                        list.Add(relId);
                        return relId;
                    }
                    id++;
                }
            }

            public void Reset(RelType relType)
            {
                if (_relIds.ContainsKey(relType))
                    _relIds.Remove(relType);
            }
        }

        #endregion Nested type: RelIdGenerator

        #region Nested type: FontInfo

        internal struct FontInfo
        {
            public XLFontValue Font;
            public uint FontId;
        };

        #endregion Nested type: FontInfo

        #region Nested type: FillInfo

        internal struct FillInfo
        {
            public XLFillValue Fill;
            public uint FillId;
        }

        #endregion Nested type: FillInfo

        #region Nested type: BorderInfo

        internal struct BorderInfo
        {
            public XLBorderValue Border;
            public uint BorderId;
        }

        #endregion Nested type: BorderInfo

        #region Nested type: NumberFormatInfo

        internal struct NumberFormatInfo
        {
            public XLNumberFormatValue NumberFormat;
            public int NumberFormatId;
        }

        #endregion Nested type: NumberFormatInfo

        #region Nested type: StyleInfo

        internal struct StyleInfo
        {
            public uint BorderId;
            public uint FillId;
            public uint FontId;
            public bool IncludeQuotePrefix;
            public int NumberFormatId;
            public XLStyleValue Style;
            public uint StyleId;
        }

        #endregion Nested type: StyleInfo

        #region Nested type: Pivot tables

        internal struct PivotTableFieldInfo
        {
            public XLDataType DataType;
            public bool MixedDataType;
            public IEnumerable<object> DistinctValues;
            public bool IsTotallyBlankField;
        }

        internal struct PivotTableInfo
        {
            public IDictionary<string, PivotTableFieldInfo> Fields;
            public Guid Guid;
        }

        #endregion Nested type: Pivot tables
    }
}
