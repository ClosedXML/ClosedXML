using System.IO;

namespace ClosedXML.Graphics
{
    internal abstract class ImageInfoReader
    {
        public bool TryGetInfo(Stream stream, out XLPictureInfo info)
        {
            info = default;
            stream.Position = 0;
            if (!CheckHeader(stream))
            {
                stream.Position = 0;
                return false;
            }

            stream.Position = 0;
            info = ReadInfo(stream);
            return true;
        }

        protected abstract bool CheckHeader(Stream stream);

        protected abstract XLPictureInfo ReadInfo(Stream stream);
    }
}
