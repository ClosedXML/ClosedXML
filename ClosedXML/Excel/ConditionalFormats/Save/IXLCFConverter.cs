using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal interface IXLCFConverter
    {
        ConditionalFormattingRule Convert(XLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context);
    }
}
