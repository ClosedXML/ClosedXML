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

        public static NonVisualDrawingProperties GetPropertiesFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
                return null;

            // Maybe we should not restrict here, and just search for all NonVisualDrawingProperties in an anchor?
            var shape = anchor.Descendants<Xdr.Picture>().Cast<OpenXmlCompositeElement>().FirstOrDefault()
                        ?? anchor.Descendants<Xdr.ConnectionShape>().Cast<OpenXmlCompositeElement>().FirstOrDefault();

            if (shape == null) return null;

            return shape
                .Descendants<Xdr.NonVisualDrawingProperties>()
                .FirstOrDefault();
        }

        public static String GetImageRelIdFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
                return null;

            var blipFill = anchor.Descendants<Xdr.BlipFill>().FirstOrDefault();
            return blipFill?.Blip?.Embed?.Value;
        }

        private static bool IsAllowedAnchor(OpenXmlElement anchor)
        {
            var allowedAnchorTypes = new Type[] { typeof(AbsoluteAnchor), typeof(OneCellAnchor), typeof(TwoCellAnchor) };
            return (allowedAnchorTypes.Any(t => t == anchor.GetType()));
        }
    }
}
