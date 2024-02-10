namespace ClosedXML.Excel;

/// <summary>
/// Specifies how to apply <see cref="XLConditionalFormatType.Top10"/> conditional formatting rule
/// on a pivot table <see cref="XLPivotConditionalFormat"/>. Avoid if possible, doesn't seem to
/// work and row/column causes Excel to repair file.
/// </summary>
/// <remarks>18.18.84 ST_Type.</remarks>
internal enum XLPivotCfRuleType
{
    All,
    Column,
    None,
    Row
}
