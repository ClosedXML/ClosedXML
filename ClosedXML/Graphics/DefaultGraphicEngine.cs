using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Graphics
{
    public class DefaultGraphicEngine : IXLGraphicEngine
    {
        /// <summary>
        /// An engine that contains an embedded Noto Sans Display font that is used for all text measurements.
        /// </summary>
        public static readonly Lazy<DefaultGraphicEngine> Embedded = new(() => new DefaultGraphicEngine(() => ReadEmbeddedFont("NotoSansDisplay.fon")));

        /// <summary>
        /// An engine that uses external font file (%SystemRoot%/Fonts/calibri.ttf) for all text measurements. If not found (non-Windows environments), an exception will be thrown.
        /// </summary>
        public static readonly Lazy<DefaultGraphicEngine> External = new(() => new DefaultGraphicEngine(() => ReadSystemFont("calibri.ttf")));

        private readonly Lazy<FontMetric> _fontMetric;
        private readonly ImageMetadataReader[] _imageReaders =
        {
            new PngMetadataReader(),
            new JpegMetadataReader(),
            new EmfMetadataReader(),
        };

        private DefaultGraphicEngine(Func<FontMetric> createFont)
        {
            _fontMetric = new Lazy<FontMetric>(createFont);
        }

        public XLPictureMetadata GetPictureMetadata(Stream stream, XLPictureFormat expectedFormat)
        {
            foreach (var imageReader in _imageReaders)
            {
                if (imageReader.TryGetDimensions(stream, out var dimensions))
                    return dimensions;
            }

            throw new ArgumentException("Unable to determine the format of the image.");
        }

        public double GetTextHeight(IXLFontBase textFont)
        {
            var fontMetric = _fontMetric.Value;
            var heightInFontUnits = fontMetric.Ascent + 2 * fontMetric.Descent;
            var pointsPerFontUnits = textFont.FontSize / fontMetric.UnitsPerEm;
            return heightInFontUnits * pointsPerFontUnits;
        }

        public double GetTextWidth(string text, IXLFontBase textFont)
        {
            var fontMetric = _fontMetric.Value;
            var widthInFontUnits = 0;
            foreach (var textCharacter in text)
                widthInFontUnits += fontMetric.GetAdvanceWidth(textCharacter);

            return widthInFontUnits * textFont.FontSize / fontMetric.UnitsPerEm;
        }

        public double GetMaxDigitWidth(IXLFontBase textFont)
        {
            var fontMetric = _fontMetric.Value;
            return fontMetric.MaxDigitWidth * textFont.FontSize / fontMetric.UnitsPerEm;
        }

        public double GetAscent(IXLFontBase fontBase)
        {
            var fontMetric = _fontMetric.Value;
            return (fontMetric.UnitsPerEm - fontMetric.Descent) / (double)fontMetric.UnitsPerEm;
        }

        public double GetDescent(IXLFontBase fontBase)
        {
            var fontMetric = _fontMetric.Value;
            return fontMetric.Descent / (double)fontMetric.UnitsPerEm;
        }

        private static FontMetric ReadEmbeddedFont(string embeddedFontName)
        {
            using var stream = typeof(DefaultGraphicEngine).Assembly.GetManifestResourceStream($"ClosedXML.Graphics.{embeddedFontName}");
            return FontMetric.LoadFromEmbedded(stream);
        }

        private static FontMetric ReadSystemFont(string fontFileName)
        {
            var fontPath = Environment.ExpandEnvironmentVariables(FormattableString.Invariant($"%SystemRoot%/Fonts/{fontFileName}"));
            using var stream = File.OpenRead(fontPath);
            return FontMetric.LoadTrueType(stream);
        }
    }
}
