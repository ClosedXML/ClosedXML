using System;

namespace ClosedXML.Excel;

public class XLExportCellOptions
{
    private int fontsSize;
    public XLColor? BackgroundColor { get; set; }

    public bool Bold { get; set; }

    public int FontSize
    {
        get => fontsSize < 6 ? 6 : fontsSize;
        set
        {
            if (value < 6)
                throw new ArgumentException("FontSize must be greater than or equal to 6.");

            fontsSize = value;
        }
    }

    public XLAlignmentHorizontalValues TextAlignHorizontal { get; set; }
    public XLAlignmentVerticalValues TextAlignVertical { get; set; }
    public XLColor? TextColor { get; set; }
}
