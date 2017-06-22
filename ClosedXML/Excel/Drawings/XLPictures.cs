using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ClosedXML.Excel.Drawings
{
    internal class XLPictures : IXLPictures, IEnumerable<XLPicture>
    {
        private readonly List<XLPicture> _pictures = new List<XLPicture>();
        private readonly XLWorksheet _worksheet;

        public XLPictures(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public int Count
        {
            [DebuggerStepThrough]
            get { return _pictures.Count; }
        }

        public IXLPicture Add(Stream stream)
        {
            var picture = new XLPicture(_worksheet, stream);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();
            return picture;
        }

        public IXLPicture Add(Stream stream, string name)
        {
            var picture = Add(stream);
            picture.Name = name;
            return picture;
        }

        public Drawings.IXLPicture Add(Stream stream, XLPictureFormat format)
        {
            var picture = new XLPicture(_worksheet, stream, format);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();
            return picture;
        }

        public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
        {
            var picture = Add(stream, format);
            picture.Name = name;
            return picture;
        }

        public IXLPicture Add(Bitmap bitmap)
        {
            var picture = new XLPicture(_worksheet, bitmap);
            _pictures.Add(picture);
            picture.Name = GetNextPictureName();
            return picture;
        }

        public IXLPicture Add(Bitmap bitmap, string name)
        {
            var picture = Add(bitmap);
            picture.Name = name;
            return picture;
        }

        public IXLPicture Add(string imageFile)
        {
            using (var bitmap = Image.FromFile(imageFile) as Bitmap)
            {
                var picture = new XLPicture(_worksheet, bitmap);
                _pictures.Add(picture);
                picture.Name = GetNextPictureName();
                return picture;
            }
        }

        public IXLPicture Add(string imageFile, string name)
        {
            var picture = Add(imageFile);
            picture.Name = name;
            return picture;
        }

        public void Delete(IXLPicture picture)
        {
            Delete(picture.Name);
        }

        public void Delete(string pictureName)
        {
            _pictures.RemoveAll(picture => picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
        }

        IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator()
        {
            return _pictures.Cast<IXLPicture>().GetEnumerator();
        }

        public IEnumerator<XLPicture> GetEnumerator()
        {
            return ((IEnumerable<XLPicture>)_pictures).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLPicture Picture(string pictureName)
        {
            IXLPicture p;

            if (TryGetPicture(pictureName, out p))
                return p;

            throw new ArgumentException($"There isn't a picture named '{pictureName}'.");
        }

        public bool TryGetPicture(string pictureName, out IXLPicture picture)
        {
            var matches = _pictures.Where(p => p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
            if (matches.Any())
            {
                picture = matches.Single();
                return true;
            }
            picture = null;
            return false;
        }

        private String GetNextPictureName()
        {
            var pictureNumber = this.Count;
            while (_pictures.Any(p => p.Name == $"Picture {pictureNumber}"))
            {
                pictureNumber++;
            }
            return $"Picture {pictureNumber}";
        }
    }
}
