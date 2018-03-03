using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFIsErrorConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            if (!cf.Style.Value.Equals(XLWorkbook.DefaultStyle.Value))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cf.Style.Value.Key];

            var formula = new Formula { Text = "ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
