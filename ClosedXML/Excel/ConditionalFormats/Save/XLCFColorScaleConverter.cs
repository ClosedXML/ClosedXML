using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            var colorScale = new ColorScale();
            for (var i = 1; i <= cf.ContentTypes.Count; i++)
            {
                var type = cf.ContentTypes[i].ToOpenXml();
                var val = cf.Values.TryGetValue(i, out var formula) ? formula?.Value : null;

                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = type };
                if (val != null)
                {
                    conditionalFormatValueObject.Val = val;
                }

                colorScale.Append(conditionalFormatValueObject);
            }

            for (var i = 1; i <= cf.Colors.Count; i++)
            {
                var xlColor = cf.Colors[i];
                var color = new Color();
                switch (xlColor.ColorType)
                {
                    case XLColorType.Color:
                        color.Rgb = xlColor.Color.ToHex();
                        break;
                    case XLColorType.Theme:
                        color.Theme = System.Convert.ToUInt32(xlColor.ThemeColor);
                        break;

                    case XLColorType.Indexed:
                        color.Indexed = System.Convert.ToUInt32(xlColor.Indexed);
                        break;
                }

                colorScale.Append(color);
            }

            conditionalFormattingRule.Append(colorScale);

            return conditionalFormattingRule;
        }
    }
}
