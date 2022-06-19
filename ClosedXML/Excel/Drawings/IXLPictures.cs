using SkiaSharp;
using System.Collections.Generic;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPictures : IEnumerable<IXLPicture>
    {
        int Count { get; }

        IXLPicture Add(Stream stream);

        IXLPicture Add(Stream stream, string name);

        IXLPicture Add(Stream stream, XLPictureFormat format);

        IXLPicture Add(Stream stream, XLPictureFormat format, string name);

        IXLPicture Add(SKCodec bitmap);

        IXLPicture Add(SKCodec bitmap, string name);

        IXLPicture Add(string imageFile);

        IXLPicture Add(string imageFile, string name);

        bool Contains(string pictureName);

        void Delete(string pictureName);

        void Delete(IXLPicture picture);

        IXLPicture Picture(string pictureName);

        bool TryGetPicture(string pictureName, out IXLPicture picture);
    }
}
