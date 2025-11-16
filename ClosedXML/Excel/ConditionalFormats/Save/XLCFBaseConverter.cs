using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal static class XLCFBaseConverter
    {
        public static ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority)
        {
            return new ConditionalFormattingRule
            {
                Type = cf.ConditionalFormatType.ToOpenXml(),
                Priority = priority,
                StopIfTrue = OpenXmlHelper.GetBooleanValue(cf.StopIfTrue, false)
            };
        }

        public static ConditionalFormattingRule ConvertWithDxf(XLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var cfRule = Convert(cf, priority);
#if STYLES_REWORK
            cfRule.FormatId = context.GetDxfId(cf.FormatValue);
#else
            var cfStyle = ((XLStyle)cf.Style).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
                cfRule.FormatId = context.GetDxfId(cfStyle);
#endif

            return cfRule;
        }
    }
}
