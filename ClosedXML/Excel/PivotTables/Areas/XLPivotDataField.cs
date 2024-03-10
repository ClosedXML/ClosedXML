namespace ClosedXML.Excel;

/// <summary>
/// A field that describes calculation of value to display in the <see cref="XLPivotAreaType.Data"/>
/// area of pivot table.
/// </summary>
internal class XLPivotDataField
{
    internal XLPivotDataField(uint field)
    {
        Field = field;
    }

    /// <summary>
    /// Custom name of the data field (e.g. <em>Sum of Sold</em>).
    /// </summary>
    internal string? DataFieldName { get; init; }

    /// <summary>
    /// Field index to <see cref="XLPivotTable.PivotFields"/>.
    /// </summary>
    /// <remarks>
    /// Unlike axis, this field index can't be <c>-2</c> for data fields. That field can't be in
    /// the data area.
    /// </remarks>
    internal uint Field { get; }

    /// <summary>
    /// An aggregation function that calculates the value to display in the data cells of pivot area.
    /// </summary>
    public XLPivotSummary Subtotal { get; init; } = XLPivotSummary.Sum;

    /// <summary>
    /// A calculation takes value calculated by <see cref="Subtotal"/> aggregation and transforms
    /// it into the final value to display to the user. The calculation might need
    /// <see cref="BaseField"/> and/or <see cref="BaseItem"/>.
    /// </summary>
    public XLPivotCalculation ShowDataAsFormat { get; init; } = XLPivotCalculation.Normal;

    /// <summary>
    /// Index to the base field (<see cref="XLPivotTable.PivotFields"/>) when
    /// <see cref="ShowDataAsFormat"/> needs a field for its calculation.
    /// </summary>
    public int BaseField { get; init; } = -1;

    /// <summary>
    /// Index to the base item of <see cref="BaseField"/> when <see cref="ShowDataAsFormat"/> needs
    /// an item for its calculation.
    /// </summary>
    public uint BaseItem { get; init; } = 1048832;

    /// <summary>
    /// Formatting to apply to the data field. If <see cref="XLPivotFormat"/> disagree, this has precedence.
    /// </summary>
    public uint? NumberFormatId { get; init; }
}
