using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

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
        long OffsetX { get; set; }
        long OffsetY { get; set; }
        long RawOffsetX { get; set; }
        long RawOffsetY { get; set; }
        bool IsAbsolute { get; set; }

        /// <summary>
        /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
        /// Default value is "jpeg"
        /// </summary>
        String Type { get; set; }

        String Name { get; set; }

        /// <summary>
        /// Get the enum representation of the Picture type.
        /// </summary>
        ImagePartType GetImagePartType();
    }
}
