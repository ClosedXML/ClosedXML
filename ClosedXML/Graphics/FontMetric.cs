using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Utils;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A simplified metrics of a font for layout. Everything is in font units. See https://stackoverflow.com/questions/13604492 for an image describing the metrics.
    /// </summary>
    internal class FontMetric
    {
        /// <summary>
        /// An extra space above ascent that is still considered a part of glyph/line.
        /// </summary>
        public int InnerLeading { get; }

        /// <summary>
        /// Ascent of glyphs above baseline. The ascent includes inner leading.
        /// </summary>
        public int Ascent { get; }

        /// <summary>
        /// Descent of glyphs below baseline.
        /// </summary>
        public int Descent { get; }

        /// <summary>
        /// A line gap, distance between bottom of upper line and top of the bottom line.
        /// </summary>
        public int ExternalLeading { get; }

        /// <summary>
        /// How many logical units are in 1 EM for this font.
        /// </summary>
        public int UnitsPerEm { get; }

        /// <summary>
        /// A maximum width of a 0-9 characters.
        /// </summary>
        public int MaxDigitWidth { get; }

        /// <summary>
        /// Get advance width for a character.
        /// </summary>
        private readonly Dictionary<char, int> _advanceWidths;

        private FontMetric(int unitsPerEm, int ascent, int descent, Dictionary<char, int> advanceWidths)
        {
            UnitsPerEm = unitsPerEm;
            Ascent = ascent;
            Descent = descent;
            _advanceWidths = advanceWidths;
            MaxDigitWidth = GetMaxDigitWidth();
        }

        public static FontMetric LoadFromEmbedded(Stream stream)
        {
            var unitsPerEm = stream.ReadS16BE();
            var ascent = stream.ReadS16BE();
            var descent = stream.ReadS16BE();
            var externalLeading = stream.ReadS16BE();
            var innerLeading = ascent + descent - unitsPerEm;

            var codepointCount = stream.ReadS32BE();
            var advanceWidths = new Dictionary<char, int>(codepointCount);
            for (var i = 0; i < codepointCount; ++i)
            {
                var codepoint = (char)stream.ReadS16BE();
                var advanceWidth = stream.ReadS16BE();
                advanceWidths[codepoint] = advanceWidth;
            }
            return new FontMetric(unitsPerEm, ascent, descent, advanceWidths);
        }

        internal int GetAdvanceWidth(char c)
        {
            return _advanceWidths.TryGetValue(c, out var advanceWidth)
                ? advanceWidth
                : UnitsPerEm;
        }

        private int GetMaxDigitWidth()
        {
            int maxDigitWidth = default;
            for (var digit = '0'; digit <= '9'; ++digit)
                maxDigitWidth = Math.Max(GetAdvanceWidth(digit), maxDigitWidth);

            return maxDigitWidth;
        }

        /// <summary>
        /// Bare bones loader of ttf fonts that is capable of loading a Calibre font layout.
        /// It expects that the font contains requires tables and that it contains a unicode
        /// encoding for BMP in format 4.
        /// </summary>
        public static FontMetric LoadTrueType(Stream s)
        {
            s.Position = 4;
            var tableCount = s.ReadU16BE();
            s.Position += 6;
            var directory = LoadTableDirectory(s, tableCount);

            var headTable = directory["head"];
            var unitsPerEm = headTable.ReadU16(s, 18);

            // Windows GDI uses an ascent/descent that has a different value from the typographic ascent/descent
            // For Calibri font, ascent + descent doesn't add up to the unitsPerEm
            var os2Table = directory["OS/2"];
            var winAscent = os2Table.ReadU16(s, 74);
            var winDescent = os2Table.ReadU16(s, 76);

            var hheaTable = directory["hhea"];
            var horizontalMetricCount = hheaTable.ReadU16(s, 34);

            var hmtxTable = directory["hmtx"];
            hmtxTable.MoveToStart(s);
            var hmtx = HmtxTable.Load(s, horizontalMetricCount);

            var cmapTable = directory["cmap"];
            cmapTable.MoveToStart(s);
            var cmap = CmapTable.Load(s);

            var advanceWidths = new Dictionary<char, int>();
            ushort codepoint = 0;
            do
            {
                if (cmap.TryGetGlyph(codepoint, out var glyphId))
                    advanceWidths.Add((char)codepoint, hmtx.GetAdvanceWidth(glyphId));
            } while (codepoint++ != 0xFFFF);

            return new FontMetric(unitsPerEm, winAscent, winDescent, advanceWidths);
        }

        private static Dictionary<string, TableRecord> LoadTableDirectory(Stream s, int tableCount)
        {
            var tableDirectory = new Dictionary<string, TableRecord>();
            for (var i = 0; i < tableCount; ++i)
            {
                var tag = Encoding.ASCII.GetString(new[] { s.ReadU8(), s.ReadU8(), s.ReadU8(), s.ReadU8() });
                s.Position += 4;
                var offset = s.ReadU32BE();
                var length = s.ReadU32BE();
                tableDirectory.Add(tag, new TableRecord(offset, length));
            }
            return tableDirectory;
        }

        public readonly struct TableRecord
        {
            private readonly UInt32 _offset;
            private readonly UInt32 _length;

            public TableRecord(UInt32 offset, UInt32 length)
            {
                _offset = offset;
                _length = length;
            }

            public ushort ReadU16(Stream s, long tableOffset)
            {
                if (tableOffset < 0 || tableOffset + 2 > _length)
                    throw new ArgumentException("Trying to read outside of a table.");
                s.Position = _offset + tableOffset;
                return s.ReadU16BE();
            }

            internal void MoveToStart(Stream s) => s.Position = _offset;
        }

        /// <summary>
        /// Horizontal metrics table.
        /// </summary>
        private class HmtxTable
        {
            private readonly IReadOnlyList<ushort> _advanceWidths;

            private HmtxTable(IReadOnlyList<ushort> advanceWidths)
                => _advanceWidths = advanceWidths;

            public ushort GetAdvanceWidth(ushort glyphId)
            {
                if (glyphId < _advanceWidths.Count)
                    return _advanceWidths[glyphId];

                // Some fonts (mostly monospace ones) reduce storage requirements by repeating the same advance width for all chars
                return _advanceWidths[_advanceWidths.Count - 1];
            }

            public static HmtxTable Load(Stream s, ushort metricCount)
            {
                var advanceWidths = new List<ushort>(metricCount);
                for (ushort glyphId = 0; glyphId < metricCount; ++glyphId)
                {
                    advanceWidths.Add(s.ReadU16BE());
                    s.Position += 2;
                }

                return new HmtxTable(advanceWidths);
            }
        }

        /// <summary>
        /// Character to glyph index mapping table.
        /// </summary>
        private class CmapTable
        {
            private readonly Format4SubTable _encodingSubTable;

            private CmapTable(Format4SubTable encodingSubTable)
                => _encodingSubTable = encodingSubTable;

            public bool TryGetGlyph(ushort codepoint, out UInt16 glyphId) => _encodingSubTable.TryGetGlyph(codepoint, out glyphId);

            public static CmapTable Load(Stream s)
            {
                s.Position += 2;
                var encodingTableCount = s.ReadU16BE();

                // Find Basic Multilingual Plane encoding
                long bmpSubtableOffset = -1;
                for (var tableIndex = 0; tableIndex < encodingTableCount; ++tableIndex)
                {
                    var platformId = s.ReadU16BE();
                    var encodingId = s.ReadU16BE();
                    var subtableOffset = s.ReadU32BE();
                    if (platformId == PlatformId.Unicode && encodingId == UnicodeEncoding.BmpOnly)
                        bmpSubtableOffset = subtableOffset;
                }
                if (bmpSubtableOffset < 0)
                    throw new ArgumentException("Font doesn't contain unicode bmp encoding.");

                var endOfSubtablesOffset = 4 + 8 * encodingTableCount;
                s.Position += bmpSubtableOffset - endOfSubtablesOffset;
                var format = s.ReadU16BE();
                if (format != 4)
                    throw new ArgumentException("Format of encoding table is not 4.");

                var subTable = Format4SubTable.Load(s);
                return new CmapTable(subTable);
            }

            private static class PlatformId
            {
                public const UInt16 Unicode = 0;
            }

            private static class UnicodeEncoding
            {
                public const UInt16 BmpOnly = 3;
            }

            private class Format4SubTable
            {
                private readonly IReadOnlyList<Segment> _segments;
                private readonly IReadOnlyList<ushort> _glyphIdArray;

                private Format4SubTable(IReadOnlyList<Segment> segments, IReadOnlyList<ushort> glyphIdArray)
                {
                    _segments = segments;
                    _glyphIdArray = glyphIdArray;
                }

                public bool TryGetGlyph(ushort codepoint, out UInt16 glyphId)
                {
                    ushort segIdx = 0;
                    // endCode of last segment is always 0xFFFF
                    while (_segments[segIdx].EndCode < codepoint)
                        segIdx++;

                    var segment = _segments[segIdx];
                    if (segment.StartCode > codepoint)
                    {
                        glyphId = 0;
                        return false;
                    }

                    if (segment.IdRangeOffset != 0)
                    {
                        // The offset is from the segment start, not glyph id array. IdRangeOffset is in bytes
                        var glyphIdx = segment.IdRangeOffset / 2 + (codepoint - segment.StartCode) - (_segments.Count - segIdx);
                        glyphId = (ushort)(_glyphIdArray[glyphIdx] + segment.IdDelta);
                    }
                    else
                    {
                        glyphId = (ushort)(codepoint + segment.IdDelta);
                    }

                    return true;
                }

                public static Format4SubTable Load(Stream s)
                {
                    ushort subTableLength = s.ReadU16BE();
                    s.Position += 2;
                    ushort segCountX2 = s.ReadU16BE();
                    s.Position += 6;
                    var segCount = segCountX2 / 2;
                    var endCodes = new ushort[segCount];
                    for (var segIdx = 0; segIdx < segCount; ++segIdx)
                        endCodes[segIdx] = s.ReadU16BE();

                    s.Position += 2;
                    var startCodes = new ushort[segCount];
                    for (var segIdx = 0; segIdx < segCount; ++segIdx)
                        startCodes[segIdx] = s.ReadU16BE();

                    var idDeltas = new ushort[segCount];
                    for (var segIdx = 0; segIdx < segCount; ++segIdx)
                        idDeltas[segIdx] = s.ReadU16BE();

                    var idRangeOffsets = new ushort[segCount];
                    for (var segIdx = 0; segIdx < segCount; ++segIdx)
                        idRangeOffsets[segIdx] = s.ReadU16BE();

                    var segments = new List<Segment>(segCount);
                    for (var segIdx = 0; segIdx < segCount; ++segIdx)
                        segments.Add(new Segment(startCodes[segIdx], endCodes[segIdx], idDeltas[segIdx], idRangeOffsets[segIdx]));

                    var headerLength = 16 + 8 * segCount;
                    var glyphIdArrayLength = (subTableLength - headerLength) / 2;
                    var glyphIdArray = new List<ushort>(glyphIdArrayLength);
                    for (var i = 0; i < glyphIdArrayLength; ++i)
                        glyphIdArray.Add(s.ReadU16BE());

                    return new Format4SubTable(segments, glyphIdArray);
                }

                private readonly struct Segment
                {
                    public Segment(ushort startCode, ushort endCode, ushort idDelta, ushort idRangeOffset)
                    {
                        StartCode = startCode;
                        EndCode = endCode;
                        IdDelta = idDelta;
                        IdRangeOffset = idRangeOffset;
                    }

                    public ushort StartCode { get; }
                    public ushort EndCode { get; }
                    public ushort IdDelta { get; }
                    public ushort IdRangeOffset { get; }
                }
            }
        }
    }
}
