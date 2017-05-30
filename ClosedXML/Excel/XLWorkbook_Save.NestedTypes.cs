using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        #region Nested type: SaveContext
        internal sealed class SaveContext
        {
            #region Private fields
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly RelIdGenerator _relIdGenerator;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Dictionary<Int32, StyleInfo> _sharedStyles;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Dictionary<Int32, NumberFormatInfo> _sharedNumberFormats;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Dictionary<IXLFont, FontInfo> _sharedFonts;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly HashSet<string> _tableNames;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private uint _tableId;
            #endregion
            #region Constructor
            public SaveContext()
            {
                _relIdGenerator = new RelIdGenerator();
                _sharedStyles = new Dictionary<Int32, StyleInfo>();
                _sharedNumberFormats = new Dictionary<int, NumberFormatInfo>();
                _sharedFonts = new Dictionary<IXLFont, FontInfo>();
                _tableNames = new HashSet<String>();
                _tableId = 0;
            }
            #endregion
            #region Public properties
            public RelIdGenerator RelIdGenerator
            {
                [DebuggerStepThrough]
                get { return _relIdGenerator; }
            }
            public Dictionary<Int32, StyleInfo> SharedStyles
            {
                [DebuggerStepThrough]
                get { return _sharedStyles; }
            }
            public Dictionary<Int32, NumberFormatInfo> SharedNumberFormats
            {
                [DebuggerStepThrough]
                get { return _sharedNumberFormats; }
            }
            public Dictionary<IXLFont, FontInfo> SharedFonts
            {
                [DebuggerStepThrough]
                get { return _sharedFonts; }
            }
            public HashSet<string> TableNames
            {
                [DebuggerStepThrough]
                get { return _tableNames; }
            }
            public uint TableId
            {
                [DebuggerStepThrough]
                get { return _tableId; }
                [DebuggerStepThrough]
                set { _tableId = value; }
            }
            public Dictionary<IXLStyle, Int32> DifferentialFormats = new Dictionary<IXLStyle, int>();
            #endregion
        }
        #endregion
        #region Nested type: RelType
        internal enum RelType
        {
            Workbook//, Worksheet
        }
        #endregion
        #region Nested type: RelIdGenerator
        internal sealed class RelIdGenerator
        {
            private readonly Dictionary<RelType, List<String>> _relIds = new Dictionary<RelType, List<String>>();

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
            public void AddValues(IEnumerable<String> values, RelType relType)
            {
                if (!_relIds.ContainsKey(relType))
                {
                    _relIds.Add(relType, new List<String>());
                }
                _relIds[relType].AddRange(values.Where(v => !_relIds[relType].Contains(v)));
            }
            public void Reset(RelType relType)
            {
                if (_relIds.ContainsKey(relType))
                    _relIds.Remove(relType);
            }
        }
        #endregion
        #region Nested type: FontInfo
        internal struct FontInfo
        {
            public UInt32 FontId;
            public XLFont Font;
        };
        #endregion
        #region Nested type: FillInfo
        internal struct FillInfo
        {
            public UInt32 FillId;
            public XLFill Fill;
        }
        #endregion
        #region Nested type: BorderInfo
        internal struct BorderInfo
        {
            public UInt32 BorderId;
            public XLBorder Border;
        }
        #endregion
        #region Nested type: NumberFormatInfo
        internal struct NumberFormatInfo
        {
            public Int32 NumberFormatId;
            public IXLNumberFormatBase NumberFormat;
        }
        #endregion
        #region Nested type: StyleInfo
        internal struct StyleInfo
        {
            public UInt32 StyleId;
            public UInt32 FontId;
            public UInt32 FillId;
            public UInt32 BorderId;
            public Int32 NumberFormatId;
            public IXLStyle Style;
        }
        #endregion
    }
}
