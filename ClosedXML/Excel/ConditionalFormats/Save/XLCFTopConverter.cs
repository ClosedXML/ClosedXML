using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFTopConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            UInt32 val = UInt32.Parse(cf.Values[1].Value);
            var conditionalFormattingRule = new ConditionalFormattingRule { FormatId = (UInt32)context.DifferentialFormats[cf.Style], Percent = cf.Percent, Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority, Rank = val, Bottom = cf.Bottom};
            return conditionalFormattingRule;
        }


    }
}
