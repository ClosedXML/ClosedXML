using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Graphics
{
    internal partial class DefaultGraphicEngine : IXLGraphicEngine
    {
        public static readonly Lazy<DefaultGraphicEngine> Instance = new();

        private readonly Lazy<FontMetric> _fontMetric = new(() => ReadFontMetric("NotoSansDisplay.fon"));
        private readonly ImageMetadataReader[] _imageReaders =
        {
            new PngMetadataReader(),
            new JpegMetadataReader(),
            new EmfMetadataReader(),
        };

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

        private static FontMetric ReadFontMetric(string embeddedFontName)
        {
            using var stream = typeof(DefaultGraphicEngine).Assembly.GetManifestResourceStream($"ClosedXML.Graphics.{embeddedFontName}");
            return new FontMetric(stream);
        }
    }
}
