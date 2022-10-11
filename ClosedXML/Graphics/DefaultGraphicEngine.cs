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
        public static readonly Lazy<DefaultGraphicEngine> Embedded = new(() => new DefaultGraphicEngine(ReadEmbeddedFont("NotoSansDisplay.fon")));

        /// <summary>
        /// An engine that uses external font file (%SystemRoot%/Fonts/calibri.ttf) for all text measurements. If not found (non-Windows environments), an exception will be thrown.
        /// </summary>
        public static readonly Lazy<DefaultGraphicEngine> External = new(() => new DefaultGraphicEngine(ReadSystemFont("calibri.ttf")));

        internal static Lazy<DefaultGraphicEngine> Instance { get; } = new(() =>
        {
            try
            {
                return External.Value;
            }
            catch
            {
                return Embedded.Value;
            }
        });

        private readonly FontMetric _fontMetric;
        private readonly ImageInfoReader[] _imageReaders =
        {
            new PngInfoReader(),
            new JpegInfoReader(),
            new EmfInfoReader(),
        };

        private DefaultGraphicEngine(FontMetric fontMetric)
        {
            _fontMetric = fontMetric;
        }

        public XLPictureInfo GetPictureInfo(Stream stream, XLPictureFormat expectedFormat)
        {
            foreach (var imageReader in _imageReaders)
            {
                if (imageReader.TryGetInfo(stream, out var dimensions))
                    return dimensions;
            }

            throw new ArgumentException("Unable to determine the format of the image.");
        }

        public double GetTextHeight(IXLFontBase textFont, double dpiY)
        {
            var fontMetric = _fontMetric;
            var heightInFontUnits = fontMetric.Ascent + 2 * fontMetric.Descent;
            var pointsPerFontUnits = textFont.FontSize / fontMetric.UnitsPerEm;
            return XLHelper.PointsToPixels(heightInFontUnits * pointsPerFontUnits, dpiY);
        }

        public double GetTextWidth(string text, IXLFontBase textFont, double dpiX)
        {
            var fontMetric = _fontMetric;
            var widthInFontUnits = 0;
            foreach (var textCharacter in text)
                widthInFontUnits += fontMetric.GetAdvanceWidth(textCharacter);

            return XLHelper.PointsToPixels(widthInFontUnits * textFont.FontSize / fontMetric.UnitsPerEm, dpiX);
        }

        public double GetMaxDigitWidth(IXLFontBase textFont, double dpiX)
        {
            var fontMetric = _fontMetric;
            return XLHelper.PointsToPixels(fontMetric.MaxDigitWidth * textFont.FontSize / fontMetric.UnitsPerEm, dpiX);
        }

        public double GetDescent(IXLFontBase fontBase, double dpiY)
        {
            var fontMetric = _fontMetric;
            return XLHelper.PointsToPixels(fontMetric.Descent * fontBase.FontSize / fontMetric.UnitsPerEm, dpiY);
        }

        private static FontMetric ReadEmbeddedFont(string embeddedFontName)
        {
            using var stream = typeof(DefaultGraphicEngine).Assembly.GetManifestResourceStream($"ClosedXML.Graphics.{embeddedFontName}");
            return FontMetric.LoadFromEmbedded(stream);
        }

        private static FontMetric ReadSystemFont(string fontFileName)
        {
            FormattableString nonExpandedFontPath = $"%SystemRoot%/Fonts/{fontFileName}";
            var fontPath = Environment.ExpandEnvironmentVariables(FormattableString.Invariant(nonExpandedFontPath));
            try
            {
                using var stream = File.OpenRead(fontPath);
                return FontMetric.LoadTrueType(stream);
            }
            catch (Exception e)
            {
                throw new ArgumentException($"Unable to get font metrics for {fontPath} ({nonExpandedFontPath}). " +
                                            $"On non-windows environments, try to use {nameof(DefaultGraphicEngine)}.{nameof(Embedded)} graphical engine " +
                                            $"or install a graphical engine from NuGet.", e);
            }
        }
    }
}
