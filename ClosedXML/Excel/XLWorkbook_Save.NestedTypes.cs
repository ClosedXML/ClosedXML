#nullable disable

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        #region Nested type: SaveContext

        internal sealed class SaveContext
        {
#if !STYLES_REWORK
            private readonly Dictionary<XLStyleValue, StyleInfo> _sharedStyles;
#endif

            public SaveContext()
            {
                RelIdGenerator = new RelIdGenerator();
                SharedFonts = new Dictionary<XLFontValue, FontInfo>();
                SavedNumberFormats = new Dictionary<string, int>();
#if !STYLES_REWORK
                DifferentialFormats = new Dictionary<XLStyleValue, int>();
                _sharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
#endif
                TableId = 0;
                TableNames = new HashSet<String>();
                PivotSourceCacheId = 0;
            }

            public RelIdGenerator RelIdGenerator { get; private set; }
            public Dictionary<XLFontValue, FontInfo> SharedFonts { get; private set; }

            /// <summary>
            /// A map of number format to a number format id for saved file. It contains all number
            /// formats from the file, all number formats used in the application (styles, pivot
            /// tables, dxf) and all predefined formats.
            /// </summary>
            public Dictionary<string, int> SavedNumberFormats { get; }

#if !STYLES_REWORK
            public IReadOnlyDictionary<XLStyleValue, StyleInfo> SharedStyles => _sharedStyles;

            public Dictionary<XLStyleValue, Int32> DifferentialFormats { get; }

#endif
            public Dictionary<XLCellFormatValue, uint> FormatMap = new();

            public uint TableId { get; set; }
            public HashSet<string> TableNames { get; private set; }

            /// <summary>
            /// A free id that can be used by the workbook to reference to a pivot cache.
            /// The <c>PivotCaches</c> element in a workbook connects the parts with pivot
            /// cache parts.
            /// </summary>
            public uint PivotSourceCacheId { get; set; }

            /// <summary>
            /// A map of shared string ids. The index is the actual index from sharedStringId and
            /// value is an mapped stringId to write to a file. The mapped stringId has no gaps
            /// between ids.
            /// </summary>
            public List<int> SstMap { get; set; }

#nullable enable
            internal int GetSharedStringId(XLCell xlCell, string text)
            {
                var sharedStringId = SstMap[xlCell.MemorySstId];
                if (sharedStringId < 0)
                {
                    throw new UnreachableException($"Unable to find text '{text}' in shared string table for cell {xlCell.SheetPoint}. " +
                                                   "That likely means reference counting is broken. As a stop-gap, try to set the " +
                                                   "text value to an unused cell to increase number of references for the text.");
                }

                return sharedStringId;
            }

            /// <summary>
            /// Get id of number format that is going to be actually saved to database.
            /// </summary>
            internal int? GetNumberFormat(XLNumberFormatValue? numberFormat)
            {
                if (numberFormat is null)
                    return null;

                return SavedNumberFormats[numberFormat.Format];
            }

            internal UInt32 GetDxfId(XLStyleValue dxf)
            {
#if STYLES_REWORK
                throw new NotImplementedException();
#else
                return (UInt32)DifferentialFormats[dxf];
#endif
            }

            internal bool TryGetDxfId(XLStyleValue dxf, out uint dxfId)
            {
#if STYLES_REWORK
                throw new NotImplementedException();
#else
                if (DifferentialFormats.TryGetValue(dxf, out var differentialFormatId))
                {
                    dxfId = (uint)differentialFormatId;
                    return true;
                }

                dxfId = default;
                return false;
#endif
            }

            internal UInt32 GetStyleId(XLStyleValue style)
            {
#if STYLES_REWORK
                return 0; // TODO: Map to real format id
#else
                return _sharedStyles[style].StyleId;
#endif
            }

            internal uint GetStyleId(XLStyleValue style, XLCellFormatValue? format)
            {
#if STYLES_REWORK
                return format is not null ? FormatMap[format] : 0;
#else
                return _sharedStyles[style].StyleId;
#endif
            }

#if !STYLES_REWORK
            internal void AddSharedStyle(XLStyleValue style, StyleInfo info)
            {
                _sharedStyles.Add(style, info);
            }

            internal void ClearSharedStyles()
            {
                _sharedStyles.Clear();
            }
#endif
#nullable disable
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
        }

        #endregion Nested type: RelIdGenerator

        #region Nested type: FontInfo

        internal struct FontInfo
        {
            public XLFontValue Font;
            public UInt32 FontId;
        };

        #endregion Nested type: FontInfo

#if !STYLES_REWORK
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
#endif
    }
}
