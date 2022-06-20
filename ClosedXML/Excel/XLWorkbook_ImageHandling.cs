using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        public static OpenXmlElement GetAnchorFromImageId(WorksheetPart worksheetPart, string relId)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<BlipFill>()
                    .Any(x => x?.Blip?.Embed?.Value.Equals(relId) ?? false)
                );

            if (!matchingAnchor.Any())
            {
                return null;
            }
            else
            {
                return matchingAnchor.First();
            }
        }

        public static OpenXmlElement GetAnchorFromImageIndex(WorksheetPart worksheetPart, int index)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<NonVisualDrawingProperties>()
                    .Any(x => x.Id.Value.Equals(Convert.ToUInt32(index + 1)))
                );

            if (!matchingAnchor.Any())
            {
                return null;
            }
            else
            {
                return matchingAnchor.First();
            }
        }

        public static NonVisualDrawingProperties GetPropertiesFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
            {
                return null;
            }

            // Maybe we should not restrict here, and just search for all NonVisualDrawingProperties in an anchor?
            var shape = anchor.Descendants<Picture>().Cast<OpenXmlCompositeElement>().FirstOrDefault()
                        ?? anchor.Descendants<ConnectionShape>().Cast<OpenXmlCompositeElement>().FirstOrDefault();

            if (shape == null)
            {
                return null;
            }

            return shape
                .Descendants<NonVisualDrawingProperties>()
                .FirstOrDefault();
        }

        public static string GetImageRelIdFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
            {
                return null;
            }

            var blipFill = anchor.Descendants<BlipFill>().FirstOrDefault();
            return blipFill?.Blip?.Embed?.Value;
        }

        private static bool IsAllowedAnchor(OpenXmlElement anchor)
        {
            var allowedAnchorTypes = new Type[] { typeof(AbsoluteAnchor), typeof(OneCellAnchor), typeof(TwoCellAnchor) };
            return allowedAnchorTypes.Any(t => t == anchor.GetType());
        }
    }
}
