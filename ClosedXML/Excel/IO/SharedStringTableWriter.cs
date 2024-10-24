using System.IO;
using System.Xml;
using ClosedXML.Utils;
using ClosedXML.Extensions;
using DocumentFormat.OpenXml.Packaging;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class SharedStringTableWriter
    {
        internal static void GenerateSharedStringTablePartContent(XLWorkbook workbook, SharedStringTablePart sharedStringTablePart,
            SaveContext context)
        {
            // Call all table headers to make sure their names are filled
            workbook.Worksheets.ForEach(w => w.Tables.ForEach(t => _ = ((XLTable)t).FieldNames.Count));

            var settings = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = XLHelper.NoBomUTF8
            };
            var partStream = sharedStringTablePart.GetStream(FileMode.Create);
            using var xml = XmlWriter.Create(partStream, settings);

            xml.WriteStartDocument();

            // Due to streaming and XLWorkbook structure, we don't know count before strings are written.
            // Attributes count and uniqueCount are optional thus are omitted.
            xml.WriteStartElement("x", "sst", Main2006SsNs);

            var sst = workbook.SharedStringTable;
            var map = sst.GetConsecutiveMap();
            context.SstMap = map;
            for (var sharedStringId = 0; sharedStringId < map.Count; ++sharedStringId)
            {
                var continuousId = map[sharedStringId];
                if (continuousId < 0)
                    continue;

                var richText = sst.GetRichText(sharedStringId);
                if (richText is not null)
                {
                    xml.WriteStartElement("si", Main2006SsNs);
                    TextSerializer.WriteRichTextElements(xml, richText, context);
                    xml.WriteEndElement(); // si
                }
                else
                {
                    xml.WriteStartElement("si", Main2006SsNs);
                    xml.WriteStartElement("t", Main2006SsNs);
                    var sharedString = sst[sharedStringId];
                    if (
                        (!sharedString.Trim().Equals(sharedString)) ||

                         (sharedString.Contains("\r") || sharedString.Contains("\n")) //also preserve whitespace in case of line breaks
                        )
                        xml.WritePreserveSpaceAttr();

                    xml.WriteString(XmlEncoder.EncodeString(sharedString));
                    xml.WriteEndElement(); // t
                    xml.WriteEndElement(); // si
                }
            }

            xml.WriteEndElement(); // SharedStringTable
            xml.Close();
        }
    }
}
