using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using System;
using System.Collections.Generic;
using System.IO;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A graphical engine that uses <c>SixLabors.ImageSharp</c> and <c>SixLabors.Fonts</c> library.
    /// </summary>
    public class SixLaborsEngine : IXLGraphicEngine
    {
        private readonly Dictionary<string, XLPictureFormat> _mimeToFormat = new()
        {
            { "image/png", XLPictureFormat.Png },
            { "image/jpeg", XLPictureFormat.Jpeg },
            { "image/gif", XLPictureFormat.Gif },
            { "image/bmp", XLPictureFormat.Bmp },
            { "image/tiff", XLPictureFormat.Tiff },
            { "image/emf", XLPictureFormat.Emf }
        };

        private readonly Configuration _configuration;

        public SixLaborsEngine()
        {
            _configuration = new Configuration { ReadOrigin = ReadOrigin.Begin };
            _configuration.ImageFormatsManager.AddImageFormatDetector(new EmfImageFormatDetector());
            _configuration.ImageFormatsManager.SetDecoder(EmfFormat.Instance, new EmfDecoder());
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
            var metrics = GetMetric(font);
            return metrics.Ascender / (double)metrics.UnitsPerEm;
        }

        private static FontMetrics GetMetric(IXLFontBase font)
        {
            if (!SystemFonts.TryGet(font.FontName, out var fontFamily))
                throw new NotImplementedException("Do some kind of fallback"); // TODO: We should 
            var font1 = fontFamily.CreateFont(11); // Size is irrelevant, but cache anyways
            var fm = font1.FontMetrics;
            return fm;
        }

        public double GetDescent(IXLFontBase font)
        {
            var metrics = GetMetric(font);
            return -metrics.Descender / (double)metrics.UnitsPerEm;
        }

        public double GetMaxDigitWidth(IXLFontBase font)
        {
            throw new NotImplementedException();
        }

        public double GetTextHeight(IXLFontBase font)
        {
            throw new NotImplementedException();
        }

        public double GetTextWidth(string text, IXLFontBase font)
        {
            throw new NotImplementedException();
        }
    }
}
