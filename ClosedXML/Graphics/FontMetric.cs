using System;
using System.Collections.Generic;
using System.IO;
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
        private readonly Dictionary<char, int> _advanceWidth;

        public FontMetric(Stream stream)
        {
            UnitsPerEm = stream.ReadS16BE();
            Ascent = stream.ReadS16BE();
            Descent = stream.ReadS16BE();
            ExternalLeading = stream.ReadS16BE();
            InnerLeading = Ascent + Descent - UnitsPerEm;

            var codepointCount = stream.ReadS32BE();
            _advanceWidth = new Dictionary<char, int>(codepointCount);
            for (var i = 0; i < codepointCount; ++i)
            {
                var codepoint = (char)stream.ReadS16BE();
                var advanceWidth = stream.ReadS16BE();
                _advanceWidth[codepoint] = advanceWidth;
            }

            MaxDigitWidth = GetMaxDigitWidth();
        }

        internal int GetAdvanceWidth(char c)
        {
            return _advanceWidth.TryGetValue(c, out var advanceWidth)
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
    }
}
