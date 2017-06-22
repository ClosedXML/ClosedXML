using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = new ConditionalFormattingRule { Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var colorScale = new ColorScale();
            for (Int32 i = 1; i <= cf.ContentTypes.Count; i++)
            {
                var type = cf.ContentTypes[i].ToOpenXml();
                var val = (cf.Values.ContainsKey(i) && cf.Values[i] != null) ? cf.Values[i].Value : null;

                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = type };
                if (val != null)
                    conditionalFormatValueObject.Val = val;

                colorScale.Append(conditionalFormatValueObject);
            }

            for (Int32 i = 1; i <= cf.Colors.Count; i++)
            {
                Color color = new Color { Rgb = cf.Colors[i].Color.ToHex() };
                colorScale.Append(color);
            }

            conditionalFormattingRule.Append(colorScale);

            return conditionalFormattingRule;
        }
    }
}
