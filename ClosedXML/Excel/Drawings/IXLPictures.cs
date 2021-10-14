using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPictures : IEnumerable<IXLPicture>
    {
        int Count { get; }

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Stream stream);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Stream stream, String name);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Stream stream, XLPictureFormat format);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Stream stream, XLPictureFormat format, String name);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Bitmap bitmap);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(Bitmap bitmap, String name);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(String imageFile);

#if NET5_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformAttribute("windows")]
#endif
        IXLPicture Add(String imageFile, String name);

        bool Contains(String pictureName);

        void Delete(String pictureName);

        void Delete(IXLPicture picture);

        IXLPicture Picture(String pictureName);

        bool TryGetPicture(String pictureName, out IXLPicture picture);
    }
}
