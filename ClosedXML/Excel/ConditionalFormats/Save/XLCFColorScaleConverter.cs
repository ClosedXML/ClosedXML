using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleConverter:IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = new ConditionalFormattingRule { Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };
            
            var colorScale = new ColorScale();
            for(Int32 i = 1; i <= cf.Values.Count; i++)
            {
                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = cf.ContentTypes[i].ToOpenXml(), Val = cf.Values[i].Value };
                colorScale.Append(conditionalFormatValueObject);
            }

            for (Int32 i = 1; i <= cf.Values.Count; i++)
            {
                Color color = new Color { Rgb = cf.Colors[i].Color.ToHex() };
                colorScale.Append(color);
            }

            conditionalFormattingRule.Append(colorScale);

            return conditionalFormattingRule;
        }
    }
}
