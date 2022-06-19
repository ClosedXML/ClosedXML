using DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Excel
{
    internal interface IXLCFConverterExtension
    {
        ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context);
    }
}
