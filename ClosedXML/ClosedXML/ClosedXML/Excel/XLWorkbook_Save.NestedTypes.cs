using System;
using System.Collections.Generic;
using System.Diagnostics;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        #region Nested type: SaveContext
        private sealed class SaveContext
        {
            #region Private fields
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly RelIdGenerator m_relIdGenerator;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Dictionary<string, uint> m_sharedStrings;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Dictionary<IXLStyle, StyleInfo> m_sharedStyles;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly HashSet<string> m_tableNames;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private uint m_tableId;
            #endregion
            #region Constructor
            public SaveContext()
            {
                m_relIdGenerator = new RelIdGenerator();
                m_sharedStrings = new Dictionary<String, UInt32>();
                m_sharedStyles = new Dictionary<IXLStyle, StyleInfo>();
                m_tableNames = new HashSet<String>();
                m_tableId = 0;
            }
            #endregion
            #region Public properties
            public RelIdGenerator RelIdGenerator
            {
                [DebuggerStepThrough]
                get { return m_relIdGenerator; }
            }
            public Dictionary<string, uint> SharedStrings
            {
                [DebuggerStepThrough]
                get { return m_sharedStrings; }
            }
            public Dictionary<IXLStyle, StyleInfo> SharedStyles
            {
                [DebuggerStepThrough]
                get { return m_sharedStyles; }
            }
            public HashSet<string> TableNames
            {
                [DebuggerStepThrough]
                get { return m_tableNames; }
            }
            public uint TableId
            {
                [DebuggerStepThrough]
                get { return m_tableId; }
                [DebuggerStepThrough]
                set { m_tableId = value; }
            }
            #endregion
        }
        #endregion
        #region Nested type: RelType
        private enum RelType
        {
            General,
            Workbook,
            Worksheet,
            Drawing
        }
        #endregion
        #region Nested type: RelIdGenerator
        private sealed class RelIdGenerator
        {
            private readonly Dictionary<RelType, List<String>> m_relIds = new Dictionary<RelType, List<String>>();
            public String GetNext(RelType relType)
            {
                if (!m_relIds.ContainsKey(relType))
                {
                    m_relIds.Add(relType, new List<String>());
                }

                Int32 id = 1;
                while (true)
                {
                    String relId = String.Format("rId{0}", id);
                    if (!m_relIds[relType].Contains(relId))
                    {
                        m_relIds[relType].Add(relId);
                        return relId;
                    }
                    id++;
                }
            }
            public void AddValues(List<String> values, RelType relType)
            {
                if (!m_relIds.ContainsKey(relType))
                {
                    m_relIds.Add(relType, new List<String>());
                }
                m_relIds[relType].AddRange(values);
            }
        }
        #endregion
        #region Nested type: FontInfo
        private struct FontInfo
        {
            public UInt32 FontId;
            public IXLFont Font;
        };
        #endregion
        #region Nested type: FillInfo
        private struct FillInfo
        {
            public UInt32 FillId;
            public IXLFill Fill;
        }
        #endregion
        #region Nested type: BorderInfo
        private struct BorderInfo
        {
            public UInt32 BorderId;
            public IXLBorder Border;
        }
        #endregion
        #region Nested type: NumberFormatInfo
        private struct NumberFormatInfo
        {
            public Int32 NumberFormatId;
            public IXLNumberFormat NumberFormat;
        }
        #endregion
        #region Nested type: StyleInfo
        private struct StyleInfo
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