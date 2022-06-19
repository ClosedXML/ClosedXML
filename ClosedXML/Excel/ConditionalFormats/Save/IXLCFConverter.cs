using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal interface IXLCFConverter
    {
        ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context);
    }
}
