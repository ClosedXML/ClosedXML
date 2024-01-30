namespace ClosedXML.Excel;

/// <summary>
/// An enum describing if <see cref="XLPivotFormat"/> applies formatting to the cells of pivot
/// table or not.
/// </summary>
/// <remarks>
/// <para>
/// [ISO-29500] 18.18.34 ST_FormatAction
/// </para>
/// <para>
/// [MS-OI29500] 2.1.761 Excel does not support the <c>Drill</c> and <c>Formula</c> values for the
/// action attribute. Therefore, neither do we, although <c>Drill</c> and <c>Formula</c> values
/// are present in the ISO <c>ST_FormatAction</c> enum.
/// </para>
/// </remarks>
internal enum XLPivotFormatAction
{
    /// <summary>
    /// No format is applied to the pivot table. This is used when formatting is cleared from
    /// already formatted cells of pivot table.
    /// </summary>
    Blank,

    /// <summary>
    /// Pivot table has formatting. This is the default value.
    /// </summary>
    Formatting,
}
