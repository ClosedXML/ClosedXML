#nullable disable

using System.Diagnostics;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A bounding box for a glyph, in pixels.
    /// </summary>
    /// <remarks>
    /// In most cases, a glyph represents a single unicode code point. In some cases,
    /// multiple code points are represented by a single glyph (emojis in most cases).
    /// Although data type is float, the actual values should be whole numbers.
    /// That best fits to Excel behavior, but there might be some cases in the future,
    /// where values can be a floats (e.g. advance could be non-pixels aligned).
    /// </remarks>
    [DebuggerDisplay("{AdvanceWidth}x{LineHeight}")]
    public readonly struct GlyphBox
    {
        /// <summary>
        /// A special glyph box that indicates a line break. Dimensions are kept at 0, so it doesn't affect any calculations.
        /// </summary>
        internal static GlyphBox LineBreak => default;

        public GlyphBox(float advanceWidth, float emSize, float descent)
        {
            AdvanceWidth = advanceWidth;
            EmSize = emSize;
            Descent = descent;
        }

        /// <summary>
        /// Advance width in px of a box for code point. Value should be whole number.
        /// </summary>
        public float AdvanceWidth { get; }

        /// <summary>
        /// Size of Em square in pixels. If em is not a square, vertical dimension of
        /// em square. Value should be whole number.
        /// </summary>
        public float EmSize { get; }

        /// <summary>
        /// Distance in px from baseline to the bottom of the box.
        /// </summary>
        /// <remarks>
        /// Descent/height is determined by font, not by codepoints
        /// of the glyph. Value should be whole number.
        /// </remarks>
        public float Descent { get; }

        internal bool IsLineBreak => AdvanceWidth == 0 && EmSize == 0 && Descent == 0;

        /// <summary>
        /// Get line width of the glyph box. It is calculated as central band with a text and
        /// a lower and an upper bands. Central band (text) has height is <c>em-square - descent</c>
        /// and the bands are <c>descent</c>.
        /// </summary>
        internal float LineHeight => EmSize + Descent;
    }
}
