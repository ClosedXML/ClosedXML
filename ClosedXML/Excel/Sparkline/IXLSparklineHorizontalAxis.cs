// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLSparklineHorizontalAxis
    {
        #region Public Properties

        XLColor Color { get; set; }

        Boolean DateAxis { get; }

        Boolean IsVisible { get; set; }

        Boolean RightToLeft { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineHorizontalAxis SetColor(XLColor value);

        IXLSparklineHorizontalAxis SetRightToLeft(Boolean value);

        IXLSparklineHorizontalAxis SetVisible(Boolean value);

        #endregion Public Methods
    }
}
