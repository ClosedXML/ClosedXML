using System.IO;

namespace ClosedXML.Graphics
{
    internal abstract class ImageMetadataReader
    {
        public bool TryGetDimensions(Stream stream, out XLPictureMetadata metadata)
        {
            metadata = default;
            stream.Position = 0;
            if (!CheckHeader(stream))
            {
                stream.Position = 0;
                return false;
            }

            stream.Position = 0;
            metadata = ReadDimensions(stream);
            return true;
        }

        protected abstract bool CheckHeader(Stream stream);

        protected abstract XLPictureMetadata ReadDimensions(Stream stream);
    }
}
