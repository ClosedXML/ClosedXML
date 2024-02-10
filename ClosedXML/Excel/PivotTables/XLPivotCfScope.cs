namespace ClosedXML.Excel;

/// <summary>
/// Defines a scope of conditional formatting applied to <see cref="XLPivotTable"/>. The scope is
/// more of a "user preference", it doesn't determine actual scope. The actual scope is determined
/// by <see cref="XLPivotConditionalFormat.Areas"/>. The scope determines what is in GUI and when
/// reapplied, it updates the <see cref="XLPivotConditionalFormat.Areas"/> according to selected
/// values.
/// </summary>
/// <remarks>18.18.67 ST_Scope</remarks>
internal enum XLPivotCfScope
{
    /// <summary>
    /// Conditional formatting is applied to selected cells. When scope is applied, CF areas are be
    /// updated to contain currently selected cells in GUI.
    /// </summary>
    SelectedCells,

    /// <summary>
    /// Conditional formatting is applied to selected data fields. When scope is applied, CF areas
    /// are be updated to contain data fields of selected cells in GUI.
    /// </summary>
    DataFields,

    /// <summary>
    /// Conditional formatting is applied to selected pivot fields intersections. When scope is
    /// applied, CF areas are be updated to contain row/column intersection of currently selected
    /// cell in GUI.
    /// </summary>
    FieldIntersections,
}
