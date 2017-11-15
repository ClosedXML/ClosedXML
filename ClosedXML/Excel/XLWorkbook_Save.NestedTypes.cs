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
                DifferentialFormats = new Dictionary<IXLStyle, int>();
                PivotTables = new Dictionary<Guid, PivotTableInfo>();
                RelIdGenerator = new RelIdGenerator();
                SharedFonts = new Dictionary<IXLFont, FontInfo>();
                SharedNumberFormats = new Dictionary<int, NumberFormatInfo>();
                SharedStyles = new Dictionary<Int32, StyleInfo>();
                TableId = 0;
                TableNames = new HashSet<String>();
            }

            public Dictionary<IXLStyle, Int32> DifferentialFormats { get; private set; }
            public IDictionary<Guid, PivotTableInfo> PivotTables { get; private set; }
            public RelIdGenerator RelIdGenerator { get; private set; }
            public Dictionary<IXLFont, FontInfo> SharedFonts { get; private set; }
            public Dictionary<Int32, NumberFormatInfo> SharedNumberFormats { get; private set; }
            public Dictionary<Int32, StyleInfo> SharedStyles { get; private set; }
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
            private readonly Dictionary<RelType, List<String>> _relIds = new Dictionary<RelType, List<String>>();

            public void AddValues(IEnumerable<String> values, RelType relType)
            {
                if (!_relIds.ContainsKey(relType))
                {
                    _relIds.Add(relType, new List<String>());
                }
                _relIds[relType].AddRange(values.Where(v => !_relIds[relType].Contains(v)));
            }

            public String GetNext()
            {
                return GetNext(RelType.Workbook);
            }

            public String GetNext(RelType relType)
            {
                if (!_relIds.ContainsKey(relType))
                {
                    _relIds.Add(relType, new List<String>());
                }

                Int32 id = _relIds[relType].Count + 1;
                while (true)
                {
                    String relId = String.Format("rId{0}", id);
                    if (!_relIds[relType].Contains(relId))
                    {
                        _relIds[relType].Add(relId);
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
            public XLFont Font;
            public UInt32 FontId;
        };

        #endregion Nested type: FontInfo

        #region Nested type: FillInfo

        internal struct FillInfo
        {
            public XLFill Fill;
            public UInt32 FillId;
        }

        #endregion Nested type: FillInfo

        #region Nested type: BorderInfo

        internal struct BorderInfo
        {
            public XLBorder Border;
            public UInt32 BorderId;
        }

        #endregion Nested type: BorderInfo

        #region Nested type: NumberFormatInfo

        internal struct NumberFormatInfo
        {
            public IXLNumberFormatBase NumberFormat;
            public Int32 NumberFormatId;
        }

        #endregion Nested type: NumberFormatInfo

        #region Nested type: StyleInfo

        internal struct StyleInfo
        {
            public UInt32 BorderId;
            public UInt32 FillId;
            public UInt32 FontId;
            public Int32 NumberFormatId;
            public IXLStyle Style;
            public UInt32 StyleId;
        }

        #endregion Nested type: StyleInfo

        #region Nested type: Pivot tables

        internal struct PivotTableFieldInfo
        {
            public XLDataType DataType;
            public Boolean MixedDataType;
            public IEnumerable<Object> DistinctValues;
        }

        internal struct PivotTableInfo
        {
            public IDictionary<String, PivotTableFieldInfo> Fields;
            public Guid Guid;
        }

        #endregion Nested type: Pivot tables
    }
}
