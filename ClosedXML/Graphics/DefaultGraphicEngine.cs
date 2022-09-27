using System;
using System.IO;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Graphics
{
    internal partial class DefaultGraphicEngine : IXLGraphicEngine
    {
        public static readonly Lazy<DefaultGraphicEngine> Instance = new();

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
    }
}
