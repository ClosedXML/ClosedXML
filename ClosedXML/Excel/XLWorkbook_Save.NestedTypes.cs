#nullable disable

using DocumentFormat.OpenXml.Packaging;
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
                TableNames = new HashSet<String>();
            }

            public Dictionary<XLStyleValue, Int32> DifferentialFormats { get; private set; }
            public IDictionary<Guid, PivotTableInfo> PivotTables { get; private set; }
            public RelIdGenerator RelIdGenerator { get; private set; }
            public Dictionary<XLFontValue, FontInfo> SharedFonts { get; private set; }
            public Dictionary<Int32, NumberFormatInfo> SharedNumberFormats { get; private set; }
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
            private readonly Dictionary<RelType, HashSet<String>> _relIds = new();

            public void AddValues(IEnumerable<String> values, RelType relType)
            {
                if (!_relIds.TryGetValue(relType, out var set))
                {
                    set = new HashSet<string>();
                    _relIds.Add(relType, set);
                }

                set.UnionWith(values);
            }

            /// <summary>
            /// Add all existing rel ids present on the parts or workbook to the generator, so they are not duplicated again.
            /// </summary>
            public void AddExistingValues(WorkbookPart workbookPart, XLWorkbook xlWorkbook)
            {
                AddValues(workbookPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                AddValues(xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !String.IsNullOrWhiteSpace(ws.RelId)).Select(ws => ws.RelId), RelType.Workbook);
                AddValues(xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !String.IsNullOrWhiteSpace(ws.LegacyDrawingId)).Select(ws => ws.LegacyDrawingId), RelType.Workbook);
                AddValues(xlWorkbook.WorksheetsInternal
                    .Cast<XLWorksheet>()
                    .SelectMany(ws => ws.Tables.Cast<XLTable>())
                    .Where(t => !String.IsNullOrWhiteSpace(t.RelId))
                    .Select(t => t.RelId), RelType.Workbook);

                foreach (var xlWorksheet in xlWorkbook.WorksheetsInternal.Cast<XLWorksheet>())
                {
                    // if the worksheet is a new one, it doesn't have RelId yet.
                    if (string.IsNullOrEmpty(xlWorksheet.RelId) || !workbookPart.TryGetPartById(xlWorksheet.RelId, out var part))
                        continue;

                    var worksheetPart = (WorksheetPart)part;
                    AddValues(worksheetPart.HyperlinkRelationships.Select(hr => hr.Id), RelType.Workbook);
                    AddValues(worksheetPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                    if (worksheetPart.DrawingsPart != null)
                        AddValues(worksheetPart.DrawingsPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                }
            }

            public String GetNext(RelType relType)
            {
                if (!_relIds.TryGetValue(relType, out var set))
                {
                    set = new HashSet<String>();
                    _relIds.Add(relType, set);
                }

                var id = set.Count + 1;
                while (true)
                {
                    var relId = String.Concat("rId", id);
                    if (!set.Contains(relId))
                    {
                        set.Add(relId);
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
            public UInt32 FontId;
        };

        #endregion Nested type: FontInfo

        #region Nested type: FillInfo

        internal struct FillInfo
        {
            public XLFillValue Fill;
            public UInt32 FillId;
        }

        #endregion Nested type: FillInfo

        #region Nested type: BorderInfo

        internal struct BorderInfo
        {
            public XLBorderValue Border;
            public UInt32 BorderId;
        }

        #endregion Nested type: BorderInfo

        #region Nested type: NumberFormatInfo

        internal struct NumberFormatInfo
        {
            public XLNumberFormatValue NumberFormat;
            public Int32 NumberFormatId;
        }

        #endregion Nested type: NumberFormatInfo

        #region Nested type: StyleInfo

        internal struct StyleInfo
        {
            public UInt32 BorderId;
            public UInt32 FillId;
            public UInt32 FontId;
            public Boolean IncludeQuotePrefix;
            public Int32 NumberFormatId;
            public XLStyleValue Style;
            public UInt32 StyleId;
        }

        #endregion Nested type: StyleInfo

        #region Nested type: Pivot tables

        internal struct PivotTableFieldInfo
        {
            public Boolean MixedDataType;
            public IReadOnlyList<XLCellValue> DistinctValues;
            public Boolean IsTotallyBlankField;
        }

        internal struct PivotTableInfo
        {
            public IDictionary<String, PivotTableFieldInfo> Fields;
            public Guid Guid;
        }

        #endregion Nested type: Pivot tables
    }
}
