using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Metadata;

namespace ClosedXML.Graphics
{
    internal class ImageInfo : IImageInfo
    {
        public PixelTypeInfo PixelType { get; }

        public int Width { get; }

        public int Height { get; }

        public ImageMetadata Metadata { get; }

        public ImageInfo(int width, int height)
        {
            Width = width;
            Height= height;
            PixelType = new PixelTypeInfo(0);
            Metadata = new ImageMetadata();
        }
    }
}
