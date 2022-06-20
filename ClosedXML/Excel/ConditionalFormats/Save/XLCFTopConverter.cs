using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFTopConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var val = uint.Parse(cf.Values[1].Value);
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            {
                conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];
            }

            conditionalFormattingRule.Percent = cf.Percent;
            conditionalFormattingRule.Rank = val;
            conditionalFormattingRule.Bottom = cf.Bottom;
            return conditionalFormattingRule;
        }
    }
}
