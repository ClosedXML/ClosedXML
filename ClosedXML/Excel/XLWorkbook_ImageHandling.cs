using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        public static OpenXmlElement GetAnchorFromImageId(WorksheetPart worksheetPart, string relId)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<Xdr.BlipFill>()
                    .Any(x => x?.Blip?.Embed?.Value.Equals(relId) ?? false)
                );

            if (!matchingAnchor.Any())
                return null;
            else
                return matchingAnchor.First();
        }

        public static OpenXmlElement GetAnchorFromImageIndex(WorksheetPart worksheetPart, Int32 index)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<Xdr.NonVisualDrawingProperties>()
                    .Any(x => x.Id.Value.Equals(Convert.ToUInt32(index + 1)))
                );

            if (!matchingAnchor.Any())
                return null;
            else
                return matchingAnchor.First();
        }

        public static NonVisualDrawingProperties GetPropertiesFromImageIndex(WorksheetPart worksheetPart, Int32 index)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            return drawingsPart.WorksheetDrawing
                .Descendants<Xdr.NonVisualDrawingProperties>()
                .FirstOrDefault(x => x.Id.Value.Equals(Convert.ToUInt32(index + 1)));
        }
    }
}
