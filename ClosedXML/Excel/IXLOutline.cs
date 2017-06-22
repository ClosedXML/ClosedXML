
namespace ClosedXML.Excel
{
    public enum XLOutlineSummaryVLocation { Top, Bottom };
    public enum XLOutlineSummaryHLocation { Left, Right };
    public interface IXLOutline
    {
        XLOutlineSummaryVLocation SummaryVLocation { get; set; }
        XLOutlineSummaryHLocation SummaryHLocation { get; set; }
    }
}
