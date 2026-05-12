#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    public interface IXLSparkline
    {
        bool IsValid { get; }

        IXLCell Location { get; set; }

        IXLRange SourceData { get; set; }

        IXLSparklineGroup SparklineGroup { get; }

        IXLSparkline SetLocation(IXLCell value);

        IXLSparkline SetSourceData(IXLRange value);
    }
}
