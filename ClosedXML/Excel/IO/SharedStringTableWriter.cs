using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO;
using System.Text;
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

            var stringId = 0;

            var newStrings = new Dictionary<String, Int32>();
            var newRichStrings = new Dictionary<IXLRichText, Int32>();

            static bool HasSharedString(XLCell c)
            {
                if (c.DataType == XLDataType.Text && c.ShareString)
                    return c.StyleValue.IncludeQuotePrefix || String.IsNullOrWhiteSpace(c.FormulaA1) && c.GetText().Length > 0;
                else
                    return false;
            }

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

            foreach (var c in workbook.Worksheets.Cast<XLWorksheet>().SelectMany(w => w.Internals.CellsCollection.GetCells(HasSharedString)))
            {
                if (c.HasRichText)
                {
                    if (newRichStrings.TryGetValue(c.GetRichText(), out int id))
                        c.SharedStringId = id;
                    else
                    {
                        xml.WriteStartElement("si", Main2006SsNs);
                        TextSerializer.WriteRichTextElements(xml, c, context);
                        xml.WriteEndElement(); // si

                        newRichStrings.Add(c.GetRichText(), stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
                else
                {
                    var value = c.Value.GetText();
                    if (newStrings.TryGetValue(value, out int id))
                        c.SharedStringId = id;
                    else
                    {
                        xml.WriteStartElement("si", Main2006SsNs);
                        xml.WriteStartElement("t", Main2006SsNs);
                        var s = value;
                        if (!s.Trim().Equals(s))
                            xml.WriteAttributeString("xml", "space", Xml1998Ns, "preserve");

                        xml.WriteString(XmlEncoder.EncodeString(s));
                        xml.WriteEndElement(); // t
                        xml.WriteEndElement(); // si

                        newStrings.Add(value, stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
            }

            xml.WriteEndElement(); // SharedStringTable
            xml.Close();
        }
    }
}
