#nullable disable

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SkiaSharp;

namespace ClosedXML.Graphics
{
    public class DefaultGraphicEngine : IXLGraphicEngine
    {
        /// <summary>
        /// Carlito is a Calibri metric compatible font. This is a version stripped of everything but metric information
        /// to keep the embedded file small. It is reasonably accurate for many alphabets (contains 2531 glyphs). It has
        /// no glyph outlines, no TTF instructions, no substitutions, glyph positioning ect. It is created from Carlito
        /// font through strip-fonts.sh script.
        /// </summary>
        private const string EmbeddedFontName = "CarlitoBare";
        private const float FontMetricSize = 16f;
        private readonly ImageInfoReader[] _imageReaders =
        {
            new PngInfoReader(),
            new JpegInfoReader(),
            new GifInfoReader(),
            new TiffInfoReader(),
            new BmpInfoReader(),
            new EmfInfoReader(),
            new WmfInfoReader(),
            new WebpInfoReader(),
            new PcxInfoReader() // Due to poor magic detection, keep last
        };

        private readonly Lazy<List<SKFont>> _fontCollection;
        private readonly string _fallbackFont;

        /// <summary>
        /// A font loaded font in the size <see cref="FontMetricSize"/>. There is no benefit in having multiple allocated instances, everything is just scaled at the moment.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, SKFont> _fonts = new();
        private readonly Func<MetricId, SKFont> _loadFont;

        /// <summary>
        /// Max digit width as a fraction of Em square. Multiply by font size to get pt size.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();

        private readonly Func<MetricId, double> _calculateMaxDigitWidth;

        /// <summary>
        /// Get a singleton instance of the engine that uses <c>Microsoft Sans Serif</c> as a fallback font.
        /// </summary>
        public static Lazy<DefaultGraphicEngine> Instance { get; } = new(() => new("Microsoft Sans Serif"));

        /// <summary>
        /// Initialize a new instance of the engine.
        /// </summary>
        /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
        public DefaultGraphicEngine(string fallbackFont)
        {
            if (string.IsNullOrWhiteSpace(fallbackFont))
                throw new ArgumentException(nameof(fallbackFont));

            var fontCollection = new List<SKFont>();

            AddEmbeddedFont(fontCollection);

             _fontCollection = new Lazy<List<SKFont>>(() => AddSystemFonts(fontCollection));
            _fallbackFont = fallbackFont;
            _loadFont = LoadFont;
            _calculateMaxDigitWidth = CalculateMaxDigitWidth;
        }

        private List<SKFont> AddSystemFonts(List<SKFont> fontCollection)
        {
            foreach (var fontFamily in SKFontManager.Default.FontFamilies)
            {
               fontCollection.Add(SKTypeface.FromFamilyName(fontFamily).ToFont());
            }

            return fontCollection;
        }

        /// <summary>
        /// Initialize a new instance of the engine. The engine will be able to use system fonts and fonts loaded from external sources.
        /// </summary>
        /// <remarks>Useful/necessary for environments without an access to filesystem.</remarks>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="useSystemFonts">Should engine try to use system fonts? If false, system fonts won't be loaded which can significantly speed up library startup.</param>
        /// <param name="fontStreams">Extra fonts that should be loaded to the engine.</param>
        private DefaultGraphicEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
        {
            if (fallbackFontStream is null)
                throw new ArgumentNullException(nameof(fallbackFontStream));

            if (fontStreams is null)
                throw new ArgumentNullException(nameof(fontStreams));

            var fontCollection = new List<SKFont>();
            AddEmbeddedFont(fontCollection);
            var fallbackFamily = SKTypeface.FromStream(fallbackFontStream).ToFont();
            fontCollection.Add(fallbackFamily);
            foreach (var fontStream in fontStreams)
                fontCollection.Add(SKTypeface.FromStream(fontStream).ToFont());
            _fontCollection = useSystemFonts
                ? new Lazy<List<SKFont>>(() => SKFontManager.Default.FontFamilies.Select(x => SKTypeface.FromFamilyName(x).ToFont()).ToList())
                : new Lazy<List<SKFont>>(() => fontCollection);
            _fallbackFont = fallbackFamily.Typeface.FamilyName;
            _loadFont = LoadFont;
            _calculateMaxDigitWidth = CalculateMaxDigitWidth;
        }

        /// <summary>
        /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
        /// It ignores all system fonts and that can lead to decrease of initialization time.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Font is determined by a name and style in the worksheet, but the font name must be mapped to a font file/stream.
        /// System fonts on Windows contain hundreds of font files that have to be checked to find the correct font
        /// file for the font name and style. That means to read hundreds of files and parse data inside them.
        /// Even though SixLabors.Fonts does this only once (lazily too) and stores data in a static variable, it is
        /// an overhead that can be avoided.
        /// </para>
        /// <para>
        /// This factory method is useful in several scenarios:
        /// <list type="bullet">
        ///   <item>Client side Blazor doesn't have access to any system fonts.</item>
        ///   <item>Worksheet contains only limited number of fonts. It might be sufficient to just load few fonts we are</item>
        /// </list>
        /// </para>
        /// </remarks>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
        public static IXLGraphicEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams)
        {
            return new DefaultGraphicEngine(fallbackFontStream, false, fontStreams);
        }

        /// <summary>
        /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
        /// It also uses system fonts.
        /// </summary>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
        public static IXLGraphicEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream,
            params Stream[] fontStreams)
        {
            return new DefaultGraphicEngine(fallbackFontStream, true, fontStreams);
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

        public double GetDescent(IXLFontBase font, double dpiY)
        {
            return GetDescent(font, dpiY, GetFont(font));
        }

        private double GetDescent(IXLFontBase font, double dpiY, SKFont skFont)
        {
            return PointsToPixels(-skFont.Metrics.Descent * font.FontSize / skFont.Typeface.UnitsPerEm, dpiY);
        }

        public double GetMaxDigitWidth(IXLFontBase fontBase, double dpiX)
        {
            var metricId = new MetricId(fontBase);
            var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
            return PointsToPixels(maxDigitWidth * fontBase.FontSize, dpiX);
        }

        public double GetTextHeight(IXLFontBase font, double dpiY)
        {
            var skFont = GetFont(font);
            var paint = new SKPaint(skFont);

            paint.TextSize = (float)font.FontSize;
            paint.GetFontMetrics(out var metrics);

            return PointsToPixels(metrics.XMax, dpiY);
        }

        public double GetTextWidth(string text, IXLFontBase fontBase, double dpiX)
        {
            var font = GetFont(fontBase);
            var paint = new SKPaint(font);
            var width = paint.MeasureText(text);
            return PointsToPixels(width / FontMetricSize * fontBase.FontSize, dpiX);
        }

        /// <inheritdoc />
        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        {
            // SixLabors.Fonts don't have a way to get a glyph representation of a cluster
            // without a TextRenderer that has unacceptable performance.
            var skFont = GetFont(font);
            var skPaint = new SKPaint(skFont);
            var advanceFu = 0f;
            for (var i = 0; i < graphemeCluster.Length; ++i)
            {
                if (!skPaint.ContainsGlyphs(graphemeCluster[i].ToString()))
                {
                    continue;
                }

                foreach (var glyphWidth in skPaint.GetGlyphWidths(graphemeCluster[i].ToString()))
                    advanceFu += glyphWidth;
            }

            var emInPx = font.FontSize / 72d * dpi.X;
            var advancePx = PointsToPixels(advanceFu * font.FontSize / skFont.Typeface.UnitsPerEm, dpi.X);
            var descentPx = GetDescent(font, dpi.Y, skFont);
            return new GlyphBox(
                (float)Math.Round(advancePx, MidpointRounding.AwayFromZero),
                (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
                (float)Math.Round(descentPx, MidpointRounding.AwayFromZero));
        }

        private SKFontMetrics GetMetrics(IXLFontBase fontBase)
        {
            var font = GetFont(fontBase);
            return font.Metrics;
        }

        private SKFont GetFont(IXLFontBase fontBase)
        {
            return GetFont(new MetricId(fontBase));
        }

        private SKFont GetFont(MetricId metricId)
        {
            return _fonts.GetOrAdd(metricId, _loadFont);
        }

        private SKFont LoadFont(MetricId metricId)
        {
            SKFont font = (_fontCollection.Value.Find(x => x.Typeface.FamilyName == metricId.Name) ??
                           _fontCollection.Value.Find(x => x.Typeface.FamilyName == _fallbackFont)) ??
                          _fontCollection.Value.First(x => x.Typeface.FamilyName == EmbeddedFontName);
            font.Size = FontMetricSize;
            return font; // Size is irrelevant for metric
        }

        private void AddEmbeddedFont(List<SKFont> fontCollection)
        {
            var assembly = Assembly.GetExecutingAssembly();
            const string resourcePath = "ClosedXML.Graphics.Fonts.CarlitoBare-{0}.ttf";

            using var regular = assembly.GetManifestResourceStream(string.Format(resourcePath, "Regular"))!;
            fontCollection.Add(SKTypeface.FromStream(regular).ToFont());

            using var bold = assembly.GetManifestResourceStream(string.Format(resourcePath, "Bold"))!;
            fontCollection.Add(SKTypeface.FromStream(bold).ToFont());

            using var italic = assembly.GetManifestResourceStream(string.Format(resourcePath, "Italic"))!;
            fontCollection.Add(SKTypeface.FromStream(italic).ToFont());

            using var boldItalic = assembly.GetManifestResourceStream(string.Format(resourcePath, "BoldItalic"))!;
            fontCollection.Add(SKTypeface.FromStream(boldItalic).ToFont());
        }

        private double CalculateMaxDigitWidth(MetricId metricId)
        {
            var font = GetFont(metricId);
            var metrics = font.Metrics;
            var maxWidth = float.MinValue;
            var skPaint = new SKPaint(font);
            for (var c = '0'; c <= '9'; ++c)
            {
                if (!skPaint.ContainsGlyphs(c.ToString()))
                {
                    continue;
                }

                var glyphAdvance = 0f;
                foreach (var glyphWidth in skPaint.GetGlyphWidths(c.ToString()))
                    glyphAdvance += glyphWidth;

                maxWidth = Math.Max(maxWidth, glyphAdvance);
            }

            return maxWidth / (double)font.Typeface.UnitsPerEm;
        }

        private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

        private readonly struct MetricId : IEquatable<MetricId>
        {
            private readonly SKFontStyle _style;

            public MetricId(IXLFontBase fontBase)
            {
                Name = fontBase.FontName;
                _style = GetFontStyle(fontBase);
            }

            public string Name { get; }

            public bool Equals(MetricId other) => Name == other.Name && _style == other._style;

            public override bool Equals(object obj) => obj is MetricId other && Equals(other);

            public override int GetHashCode() => (Name.GetHashCode() * 397) ^ StyleToInt(_style);

            private int StyleToInt(SKFontStyle style)
            {
                if (style == SKFontStyle.BoldItalic)
                {
                    return 1;
                }

                if (style == SKFontStyle.Bold)
                {
                    return 2;
                }

                if (style == SKFontStyle.Italic)
                {
                    return 3;
                }

                if (style == SKFontStyle.Normal)
                {
                    return 4;
                }

                return 0;
            }

            private static SKFontStyle GetFontStyle(IXLFontBase fontBase)
            {
                return fontBase switch
                {
                    { Bold: true, Italic: true } => SKFontStyle.BoldItalic,
                    { Bold: true } => SKFontStyle.Bold,
                    { Italic: true } => SKFontStyle.Italic,
                    _ => SKFontStyle.Normal
                };
            }
        }
    }
}
