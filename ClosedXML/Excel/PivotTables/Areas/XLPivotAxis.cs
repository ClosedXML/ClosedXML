namespace ClosedXML.Excel;

/// <summary>
/// Describes an axis of a pivot table. Used to determine which areas should be styled through
/// <see cref="XLPivotFormat.PivotArea"/>.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.18.1 ST_Axis(PivotTable Axis).
/// </remarks>
internal enum XLPivotAxis
{
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}
