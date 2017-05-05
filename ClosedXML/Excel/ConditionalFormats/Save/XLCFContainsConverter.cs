﻿using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFContainsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            conditionalFormattingRule.FormatId = (UInt32) context.DifferentialFormats[cf.Style];
            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.ContainsText;
            conditionalFormattingRule.Text = val;
        
            var formula = new Formula { Text = "NOT(ISERROR(SEARCH(\"" + val + "\"," + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")))" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }


    }
}
