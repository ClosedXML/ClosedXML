using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPictures : IEnumerable<IXLPicture>
    {
        int Count { get; }

        IXLPicture Add(Stream stream);

        IXLPicture Add(Stream stream, String name);

        IXLPicture Add(Stream stream, XLPictureFormat format);

        IXLPicture Add(Stream stream, XLPictureFormat format, String name);

        IXLPicture Add(Bitmap bitmap);

        IXLPicture Add(Bitmap bitmap, String name);

        IXLPicture Add(String imageFile);

        IXLPicture Add(String imageFile, String name);

        bool Contains(String pictureName);

        void Delete(String pictureName);

        void Delete(IXLPicture picture);

        IXLPicture Picture(String pictureName);

        bool TryGetPicture(String pictureName, out IXLPicture picture);
    }
}
