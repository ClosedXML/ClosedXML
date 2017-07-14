using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarConverter:IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = new ConditionalFormattingRule { Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var dataBar = new DataBar {ShowValue = !cf.ShowBarOnly};
            var conditionalFormatValueObject1 = new ConditionalFormatValueObject { Type = cf.ContentTypes[1].ToOpenXml()};
            if (cf.Values.Count >= 1) conditionalFormatValueObject1.Val = cf.Values[1].Value;

            var conditionalFormatValueObject2 = new ConditionalFormatValueObject { Type = cf.ContentTypes[2].ToOpenXml()};
            if (cf.Values.Count >= 2) conditionalFormatValueObject2.Val = cf.Values[2].Value;

            var color = new Color { Rgb = cf.Colors[1].Color.ToHex() };

            dataBar.Append(conditionalFormatValueObject1);
            dataBar.Append(conditionalFormatValueObject2);
            dataBar.Append(color);

            conditionalFormattingRule.Append(dataBar);

            return conditionalFormattingRule;
        }
    }
}
