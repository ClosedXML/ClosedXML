using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFContainsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = new ConditionalFormattingRule { FormatId = (UInt32)context.DifferentialFormats[cf.Style], Operator = ConditionalFormattingOperatorValues.ContainsText, Text = val, Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var formula = new Formula { Text = "NOT(ISERROR(SEARCH(\"" + val + "\"," + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")))" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }


    }
}
