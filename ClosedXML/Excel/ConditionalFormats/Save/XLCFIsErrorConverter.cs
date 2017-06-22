using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFIsErrorConverter : IXLCFConverter
    {

        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = new ConditionalFormattingRule { FormatId = (UInt32)context.DifferentialFormats[cf.Style], Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var formula = new Formula { Text = "ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }

    }
}
