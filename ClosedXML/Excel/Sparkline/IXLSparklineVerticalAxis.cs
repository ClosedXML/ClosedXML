// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public interface IXLSparklineVerticalAxis
    {
        #region Public Properties

        double? ManualMax { get; set; }

        double? ManualMin { get; set; }

        XLSparklineAxisMinMax MaxAxisType { get; set; }

        XLSparklineAxisMinMax MinAxisType { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineVerticalAxis SetManualMax(double? value);

        IXLSparklineVerticalAxis SetManualMin(double? value);

        IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax value);

        IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax value);

        #endregion Public Methods
    }
}
