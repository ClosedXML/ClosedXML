#nullable disable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel.Formatting;
using DocumentFormat.OpenXml.Packaging;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        #region Nested type: SaveContext

        internal sealed class SaveContext
        {
            public SaveContext()
            {
                RelIdGenerator = new RelIdGenerator();
                TableId = 0;
                TableNames = new HashSet<String>();
                PivotSourceCacheId = 0;
            }

            public RelIdGenerator RelIdGenerator { get; private set; }

            /// <summary>
            /// A map of number format to a number format id for saved file. It contains all number
            /// formats from the file, all number formats used in the application (styles, pivot
            /// tables, dxf) and all predefined formats.
            /// </summary>
            internal Dictionary<string, int> NumberFormatMap = new();

            internal Dictionary<XLFontFormatValue, int> FontMap = new();

            internal Dictionary<XLCellFormatValue, uint> FormatMap = new();

            internal Dictionary<XLDxfValue, uint> DxfMap = new();

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
                    throw new UnreachableException($"Unable to find text '{text}' in shared string table for cell {xlCell.Point}. " +
                                                   "That likely means reference counting is broken. As a stop-gap, try to set the " +
                                                   "text value to an unused cell to increase number of references for the text.");
                }

                return sharedStringId;
            }

            /// <summary>
            /// Get id of number format that is going to be actually saved to database.
            /// </summary>
            internal int? GetNumberFormatId(string? numberFormat)
            {
                if (numberFormat is null)
                    return null;

                return NumberFormatMap[numberFormat];
            }

            internal int GetFontId(XLFontFormatValue font)
            {
                return FontMap[font];
            }

            internal uint GetDxfId(XLDxfValue dxf)
            {
                return DxfMap[dxf];
            }

            internal bool TryGetDxfId(XLCellFormatValue dxf, out uint dxfId)
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

            internal uint GetStyleId(XLCellFormatValue? format)
            {
#if STYLES_REWORK
                return format is not null ? FormatMap[format] : 0;
#else
                return _sharedStyles[style].StyleId;
#endif
            }
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
