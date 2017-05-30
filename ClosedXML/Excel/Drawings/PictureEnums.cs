using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXML.Excel.Drawings
{
    public enum XLPictureFormat
    {
        Bmp = 0,
        Gif = 1,
        Png = 2,
        Tiff = 3,
        Icon = 4,
        Pcx = 5,
        Jpeg = 6,
        Emf = 7,
        Wmf = 8
    }

    public enum XLPicturePlacement
    {
        MoveAndSize = 0,
        Move = 1,
        FreeFloating = 2
    }
}
