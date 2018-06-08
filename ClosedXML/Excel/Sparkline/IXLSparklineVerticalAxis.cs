// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLSparklineVerticalAxis
    {
        #region Public Properties

        Double? ManualMax { get; set; }

        Double? ManualMin { get; set; }

        XLSparklineAxisMinMax MaxAxisType { get; set; }

        XLSparklineAxisMinMax MinAxisType { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineVerticalAxis SetManualMax(Double? value);

        IXLSparklineVerticalAxis SetManualMin(Double? value);

        IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax value);

        IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax value);

        #endregion Public Methods
    }
}
