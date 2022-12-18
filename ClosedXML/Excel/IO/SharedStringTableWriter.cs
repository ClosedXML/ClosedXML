using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class SharedStringTableWriter
    {
        internal static void GenerateSharedStringTablePartContent(XLWorkbook workbook, SharedStringTablePart sharedStringTablePart,
            SaveContext context)
        {
            // Call all table headers to make sure their names are filled
            var x = 0;
            workbook.Worksheets.ForEach(w => w.Tables.ForEach(t => x = (t as XLTable).FieldNames.Count));

            sharedStringTablePart.SharedStringTable = new SharedStringTable { Count = 0, UniqueCount = 0 };

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

            foreach (var c in workbook.Worksheets.Cast<XLWorksheet>().SelectMany(w => w.Internals.CellsCollection.GetCells(HasSharedString)))
            {
                if (c.HasRichText)
                {
                    if (newRichStrings.TryGetValue(c.GetRichText(), out int id))
                        c.SharedStringId = id;
                    else
                    {
                        var sharedStringItem = new SharedStringItem();
                        TextSerializer.PopulatedRichTextElements(sharedStringItem, c, context);

                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

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
                        var s = value;
                        var sharedStringItem = new SharedStringItem();
                        var text = new Text { Text = XmlEncoder.EncodeString(s) };
                        if (!s.Trim().Equals(s))
                            text.Space = SpaceProcessingModeValues.Preserve;
                        sharedStringItem.Append(text);
                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                        newStrings.Add(value, stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
            }
        }
    }
}
