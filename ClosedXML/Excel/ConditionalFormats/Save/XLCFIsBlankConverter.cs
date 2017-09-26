using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFIsBlankConverter : IXLCFConverter
    {

        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cf.Style];

            var formula = new Formula { Text = "LEN(TRIM(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))=0" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }

    }
}
