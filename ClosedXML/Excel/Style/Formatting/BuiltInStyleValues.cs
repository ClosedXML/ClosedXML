namespace ClosedXML.Excel.Formatting;

/// <summary>
/// List of built-in cell styles. The built-in style RowLevel_* is expanded from builtIn 1
/// to 101-107, depending on the outline level. Same is true for ColLevel_* that is expanded
/// from 2 to 201-207.
/// </summary>
internal enum BuiltInStyleValues
{
    /// <summary>
    /// Normal
    /// </summary>
    Normal = 0,

    /// <summary>
    /// RowLevel_1
    /// </summary>
    RowLevel1 = 101,

    /// <summary>
    /// RowLevel_2
    /// </summary>
    RowLevel2 = 102,

    /// <summary>
    /// RowLevel_3
    /// </summary>
    RowLevel3 = 103,

    /// <summary>
    /// RowLevel_4
    /// </summary>
    RowLevel4 = 104,

    /// <summary>
    /// RowLevel_5
    /// </summary>
    RowLevel5 = 105,

    /// <summary>
    /// RowLevel_6
    /// </summary>
    RowLevel6 = 106,

    /// <summary>
    /// RowLevel_7
    /// </summary>
    RowLevel7 = 107,

    /// <summary>
    /// ColLevel_1
    /// </summary>
    ColumnLevel1 = 201,

    /// <summary>
    /// ColLevel_2
    /// </summary>
    ColumnLevel2 = 202,

    /// <summary>
    /// ColLevel_3
    /// </summary>
    ColumnLevel3 = 203,

    /// <summary>
    /// ColLevel_4
    /// </summary>
    ColumnLevel4 = 204,

    /// <summary>
    /// ColLevel_5
    /// </summary>
    ColumnLevel5 = 205,

    /// <summary>
    /// ColLevel_6
    /// </summary>
    ColumnLevel6 = 206,

    /// <summary>
    /// ColLevel_7
    /// </summary>
    ColumnLevel7 = 207,

    /// <summary>
    /// Comma
    /// </summary>
    Comma = 3,

    /// <summary>
    /// Currency
    /// </summary>
    Currency = 4,

    /// <summary>
    /// Percent
    /// </summary>
    Percent = 5,

    /// <summary>
    /// Comma [0]
    /// </summary>
    Comma0 = 6,

    /// <summary>
    /// Currency [0]
    /// </summary>
    Currency0 = 7,

    /// <summary>
    /// Hyperlink
    /// </summary>
    Hyperlink = 8,

    /// <summary>
    /// Followed Hyperlink
    /// </summary>
    FollowedHyperlink = 9,

    /// <summary>
    /// Note
    /// </summary>
    Note = 10,

    /// <summary>
    /// Warning Text
    /// </summary>
    WarningText = 11,

    /// <summary>
    /// Emphasis 1
    /// </summary>
    /// <remarks>Deprecated style.</remarks>
    Emphasis1 = 12,

    /// <summary>
    /// Emphasis 2
    /// </summary>
    /// <remarks>Deprecated style.</remarks>
    Emphasis2 = 13,

    /// <summary>
    /// Emphasis 3
    /// </summary>
    /// <remarks>Deprecated style.</remarks>
    Emphasis3 = 14,

    /// <summary>
    /// Title
    /// </summary>
    Title = 15,

    /// <summary>
    /// Heading 1
    /// </summary>
    Heading1 = 16,

    /// <summary>
    /// Heading 2
    /// </summary>
    Heading2 = 17,

    /// <summary>
    /// Heading 3
    /// </summary>
    Heading3 = 18,

    /// <summary>
    /// Heading 4
    /// </summary>
    Heading4 = 19,

    /// <summary>
    /// Input
    /// </summary>
    Input = 20,

    /// <summary>
    /// Output
    /// </summary>
    Output = 21,

    /// <summary>
    /// Calculation
    /// </summary>
    Calculation = 22,

    /// <summary>
    /// Check Cell
    /// </summary>
    CheckCell = 23,

    /// <summary>
    /// Linked Cell
    /// </summary>
    LinkedCell = 24,

    /// <summary>
    /// Total
    /// </summary>
    Total = 25,

    /// <summary>
    /// Good
    /// </summary>
    Good = 26,

    /// <summary>
    /// Bad
    /// </summary>
    Bad = 27,

    /// <summary>
    /// Neutral
    /// </summary>
    Neutral = 28,

    /// <summary>
    /// Accent1
    /// </summary>
    Accent1 = 29,

    /// <summary>
    /// 20% - Accent1
    /// </summary>
    Accent1At20Percent = 30,

    /// <summary>
    /// 40% - Accent1
    /// </summary>
    Accent1At40Percent = 31,

    /// <summary>
    /// 60% - Accent1
    /// </summary>
    Accent1At60Percent = 32,

    /// <summary>
    /// Accent2
    /// </summary>
    Accent2 = 33,

    /// <summary>
    /// 20% - Accent2
    /// </summary>
    Accent2At20Percent = 34,

    /// <summary>
    /// 40% - Accent2
    /// </summary>
    Accent2At40Percent = 35,

    /// <summary>
    /// 60% - Accent2
    /// </summary>
    Accent2At60Percent = 36,

    /// <summary>
    /// Accent3
    /// </summary>
    Accent3 = 37,

    /// <summary>
    /// 20% - Accent3
    /// </summary>
    Accent3At20Percent = 38,

    /// <summary>
    /// 40% - Accent3
    /// </summary>
    Accent3At40Percent = 39,

    /// <summary>
    /// 60% - Accent3
    /// </summary>
    Accent3At60Percent = 40,

    /// <summary>
    /// Accent4
    /// </summary>
    Accent4 = 41,

    /// <summary>
    /// 20% - Accent4
    /// </summary>
    Accent4At20Percent = 42,

    /// <summary>
    /// 40% - Accent4
    /// </summary>
    Accent4At40Percent = 43,

    /// <summary>
    /// 60% - Accent4
    /// </summary>
    Accent4At60Percent = 44,

    /// <summary>
    /// Accent5
    /// </summary>
    Accent5 = 45,

    /// <summary>
    /// 20% - Accent5
    /// </summary>
    Accent5At20Percent = 46,

    /// <summary>
    /// 40% - Accent5
    /// </summary>
    Accent5At40Percent = 47,

    /// <summary>
    /// 60% - Accent5
    /// </summary>
    Accent5At60Percent = 48,

    /// <summary>
    /// Accent6
    /// </summary>
    Accent6 = 49,

    /// <summary>
    /// 20% - Accent6
    /// </summary>
    Accent6At20Percent = 50,

    /// <summary>
    /// 40% - Accent6
    /// </summary>
    Accent6At40Percent = 51,

    /// <summary>
    /// 60% - Accent6
    /// </summary>
    Accent6At60Percent = 52,

    /// <summary>
    /// Explanatory Text
    /// </summary>
    ExplanatoryText = 53,
}
