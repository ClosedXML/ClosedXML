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
        private readonly Dictionary<string, IXLPicture> _pictures = new Dictionary<string, IXLPicture>(StringComparer.OrdinalIgnoreCase);
        private readonly XLWorksheet _worksheet;

        public XLPictures(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            Deleted = new HashSet<String>();
        }

        public int Count
        {
            [DebuggerStepThrough]
            get { return _pictures.Count; }
        }

        internal ICollection<String> Deleted { get; private set; }

        public IXLPicture Add(Stream stream)
        {
            return AddInternal(new XLPicture(_worksheet, stream));
        }

        public IXLPicture Add(Stream stream, string name)
        {
            return AddInternal(new XLPicture(_worksheet, stream), name);
        }

        public IXLPicture Add(Stream stream, XLPictureFormat format)
        {
            return AddInternal(new XLPicture(_worksheet, stream, format));
        }

        public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
        {
            return AddInternal(new XLPicture(_worksheet, stream, format), name);
        }

        public IXLPicture Add(Bitmap bitmap)
        {
            return AddInternal(new XLPicture(_worksheet, bitmap));
        }

        public IXLPicture Add(Bitmap bitmap, string name)
        {
            return AddInternal(new XLPicture(_worksheet, bitmap), name);
        }

        public IXLPicture Add(string imageFile)
        {
            return Add(imageFile, null);
        }

        public IXLPicture Add(string imageFile, string name)
        {
            using (var fs = File.Open(imageFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return AddInternal(new XLPicture(_worksheet, fs), name);
            }
        }

        internal IXLPicture Add(Stream stream, string name, int Id)
        {
            if (_pictures.ContainsKey(name))
                name = GetNextPictureName();

            var picture = Add(stream, name) as XLPicture;

            if (!_worksheet.Pictures.Any(p => p.Id == Id))
                picture.Id = Id; // Use the value from the file only if it is not used already

            return picture;
        }

        private IXLPicture AddInternal(IXLPicture picture, string pictureName = null)
        {
            if (pictureName == null)
                pictureName = GetNextPictureName();
            picture.Name = pictureName;
            _pictures.Add(pictureName, picture);
            return picture;
        }

        public bool Contains(string pictureName)
        {
            return _pictures.ContainsKey(pictureName);
        }

        public void Delete(IXLPicture picture)
        {
            _pictures.Remove(picture.Name);
            
            var relId = (picture as XLPicture)?.RelId;
            if (!string.IsNullOrEmpty(relId))
                Deleted.Add(relId);

            picture.Dispose();
        }

        public void Delete(string pictureName)
        {
            if (!_pictures.ContainsKey(pictureName))
                throw new ArgumentOutOfRangeException($"Picture with name '{pictureName}' does not exist");
            Delete(_pictures[pictureName]);
        }

        IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator()
        {
            return _pictures.Values.GetEnumerator();
        }

        public IEnumerator<XLPicture> GetEnumerator()
        {
            return _pictures.Values.Cast<XLPicture>().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLPicture Picture(string pictureName)
        {
            if (TryGetPicture(pictureName, out IXLPicture p))
                return p;

            throw new ArgumentException($"There isn't a picture named '{pictureName}'.");
        }

        public bool TryGetPicture(string pictureName, out IXLPicture picture)
        {
            return _pictures.TryGetValue(pictureName, out picture);
        }

        private String GetNextPictureName()
        {
            var pictureNumber = this.Count + 1;
            while (_pictures.ContainsKey($"Picture {pictureNumber}"))
            {
                pictureNumber++;
            }
            return $"Picture {pictureNumber}";
        }

        internal void Rename(string oldName, string newName)
        {
            if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
                return;

            if (!_pictures.ContainsKey(oldName))
                return;

            var p = _pictures[oldName];

            if (TryGetPicture(newName, out var otherPicture) && p != otherPicture)
                throw new ArgumentException($"The picture name '{newName}' already exists.");

            _pictures.Remove(oldName);
            _pictures.Add(newName, p);
        }
    }
}
