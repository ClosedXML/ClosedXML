using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Bmp;
using SixLabors.ImageSharp.Formats.Gif;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Formats.Tiff;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A graphical engine that uses <c>SixLabors.ImageSharp</c> and <c>SixLabors.Fonts</c> library.
    /// </summary>
    public class SixLaborsEngine : IXLGraphicEngine
    {
        private const float fontMetricSize = 16f;
        private readonly static Lazy<SixLaborsEngine> _instance = new(() => new SixLaborsEngine());
        private static readonly Dictionary<string, XLPictureFormat> _mimeToFormat = new()
        {
            { "image/png", XLPictureFormat.Png },
            { "image/jpeg", XLPictureFormat.Jpeg },
            { "image/gif", XLPictureFormat.Gif },
            { "image/bmp", XLPictureFormat.Bmp },
            { "image/tiff", XLPictureFormat.Tiff },
            { "image/emf", XLPictureFormat.Emf }
        };

        private readonly Configuration _configuration;
        private readonly string _fallbackFont;

        /// <summary>
        /// A font loaded font in the size <see cref="fontMetricSize"/>. There is no benefit in having multiple allocated instances, everything is just scaled at the moment.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, Font> _fonts = new();
        private readonly Func<MetricId, Font> _loadFont;

        /// <summary>
        /// Max digit width as a fraction of Em square. Multiply by font size to get pt size.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();
        private readonly Func<MetricId, double> _calculateMaxDigitWidth;

        /// <summary>
        /// Initialize a new instance of the engine with the fallback font <c>Microsoft Sans Serif</c>.
        /// </summary>
        public SixLaborsEngine() : this("Microsoft Sans Serif")
        {
        }

        /// <summary>
        /// Get a singleton instance of the engine.
        /// </summary>
        public static SixLaborsEngine Instance => _instance.Value;

        /// <summary>
        /// Initialize a new instance of the engine.
        /// </summary>
        /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
        public SixLaborsEngine(string fallbackFont)
        {
            if (string.IsNullOrWhiteSpace(fallbackFont))
                throw new ArgumentException(nameof(fallbackFont));

            _configuration = new Configuration(
                new PngConfigurationModule(),
                new JpegConfigurationModule(),
                new GifConfigurationModule(),
                new BmpConfigurationModule(),
                new TiffConfigurationModule())
            {
                ReadOrigin = ReadOrigin.Begin
            };
            _configuration.ImageFormatsManager.AddImageFormatDetector(new EmfImageFormatDetector());
            _configuration.ImageFormatsManager.SetDecoder(EmfFormat.Instance, new EmfDecoder());
            _fallbackFont = fallbackFont;
            _loadFont = LoadFont;
            _calculateMaxDigitWidth = CalculateMaxDigitWidth;
        }

        public XLPictureMetadata GetPictureMetadata(Stream imageStream, XLPictureFormat expectedFormat)
        {
            var imageFormat = Image.DetectFormat(_configuration, imageStream);
            if (imageFormat is null)
                throw new ArgumentException("Unable to identity image format.");

            if (!_mimeToFormat.TryGetValue(imageFormat.DefaultMimeType, out var pictureFormat))
                pictureFormat = XLPictureFormat.Unknown;

            var imageInfo = Image.Identify(_configuration, imageStream);
            if (imageInfo is null)
                throw new ArgumentException("Unable to read image info.");

            if (imageFormat == EmfFormat.Instance)
            {
                var metadata = imageInfo.Metadata.GetFormatMetadata(EmfFormat.Instance);
                return new XLPictureMetadata(pictureFormat,
                    new System.Drawing.Size(imageInfo.Width, imageInfo.Height), new System.Drawing.Size(metadata.Frame.Width, metadata.Frame.Height),
                    imageInfo.Metadata.HorizontalResolution, imageInfo.Metadata.VerticalResolution);
            }

            return new XLPictureMetadata(pictureFormat,
                new System.Drawing.Size(imageInfo.Width, imageInfo.Height), System.Drawing.Size.Empty,
                imageInfo.Metadata.HorizontalResolution, imageInfo.Metadata.VerticalResolution);
        }

        public double GetAscent(IXLFontBase font)
        {
            var metrics = GetMetrics(font);
            return metrics.Ascender * font.FontSize / metrics.UnitsPerEm;
        }

        public double GetDescent(IXLFontBase font)
        {
            var metrics = GetMetrics(font);
            return -metrics.Descender * font.FontSize / metrics.UnitsPerEm;
        }

        public double GetMaxDigitWidth(IXLFontBase fontBase)
        {
            var metricId = new MetricId(fontBase);
            var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
            return maxDigitWidth * fontBase.FontSize;
        }

        public double GetTextHeight(IXLFontBase font)
        {
            var metrics = GetMetrics(font);
            return (metrics.Ascender - 2 * metrics.Descender) * font.FontSize / metrics.UnitsPerEm;
        }

        public double GetTextWidth(string text, IXLFontBase fontBase)
        {
            var font = GetFont(fontBase);
            var dimensionsPx = TextMeasurer.Measure(text, new TextOptions(font)
            {
                Dpi = 72, // Normalize DPI, so 1px is 1pt
                KerningMode = KerningMode.None
            });
            return dimensionsPx.Width / fontMetricSize * fontBase.FontSize;
        }

        private FontMetrics GetMetrics(IXLFontBase fontBase)
        {
            var font = GetFont(fontBase);
            return font.FontMetrics;
        }

        private Font GetFont(IXLFontBase fontBase)
        {
            return GetFont(new MetricId(fontBase));
        }

        private Font GetFont(MetricId metricId)
        {
            return _fonts.GetOrAdd(metricId, _loadFont);
        }

        private Font LoadFont(MetricId metricId)
        {
            if (!SystemFonts.TryGet(metricId.Name, out var fontFamily) &&
                !SystemFonts.TryGet(_fallbackFont, out fontFamily))
                throw new ArgumentException($"Unable to find font {metricId.Name} or fallback font {_fallbackFont}.");

            return fontFamily.CreateFont(fontMetricSize); // Size is irrelevant for metric
        }

        private double CalculateMaxDigitWidth(MetricId metricId)
        {
            var font = GetFont(metricId);
            var metrics = font.FontMetrics;
            var maxWidth = int.MinValue;
            for (var c = '0'; c <= '9'; ++c)
            {
                var glyphMetrics = metrics.GetGlyphMetrics(new SixLabors.Fonts.Unicode.CodePoint(c), ColorFontSupport.None);
                var glyphAdvance = 0;
                foreach (var glyphMetric in glyphMetrics)
                    glyphAdvance += glyphMetric.AdvanceWidth;

                maxWidth = Math.Max(maxWidth, glyphAdvance);
            }
            return maxWidth / (double)metrics.UnitsPerEm;
        }

        private readonly struct MetricId : IEquatable<MetricId>
        {
            public MetricId(IXLFontBase fontBase)
            {
                Name = fontBase.FontName;
                Style = GetFontStyle(fontBase);
            }

            public string Name { get; }

            public FontStyle Style { get; }

            public bool Equals(MetricId other) => Name == other.Name && Style == other.Style;

            public override bool Equals(object obj) => obj is MetricId other && Equals(other);

            public override int GetHashCode() => (Name.GetHashCode() * 397) ^ (int)Style;

            private static FontStyle GetFontStyle(IXLFontBase fontBase)
            {
                return fontBase switch
                {
                    { Bold: true, Italic: true } => FontStyle.BoldItalic,
                    { Bold: true } => FontStyle.Bold,
                    { Italic: true } => FontStyle.Italic,
                    _ => FontStyle.Regular
                };
            }
        }
    }
}
