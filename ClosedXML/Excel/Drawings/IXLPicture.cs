using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLPicture
    {
        Stream ImageStream { get; set; }

        List<IXLMarker> GetMarkers();
        void AddMarker(IXLMarker marker);

        long MaxHeight { get; set; }
        long MaxWidth { get; set; }
        long Width { get; set; }
        long Height { get; set; }
        long PaddingX { get; set; }
        long PaddingY { get; set; }
        long EMUOffsetX { get; set; }
        long EMUOffsetY { get; set; }

        String Name { get; set; }
    }
}
