// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public interface IXLSparklineHorizontalAxis
    {
        #region Public Properties

        XLColor Color { get; set; }

        bool DateAxis { get; }

        bool IsVisible { get; set; }

        bool RightToLeft { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineHorizontalAxis SetColor(XLColor value);

        IXLSparklineHorizontalAxis SetRightToLeft(bool value);

        IXLSparklineHorizontalAxis SetVisible(bool value);

        #endregion Public Methods
    }
}
