using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// Specification of conditional formatting of a pivot table.
/// </summary>
internal class XLPivotConditionalFormat
{
    private readonly List<XLPivotArea> _area = new();

    internal XLPivotConditionalFormat(XLConditionalFormat format)
    {
        Format = format;
    }

    /// <summary>
    /// An option to display in GUI on how to update <see cref="Areas"/>.
    /// </summary>
    internal XLPivotCfScope Scope { get; init; } = XLPivotCfScope.SelectedCells;

    /// <summary>
    /// A rule that determines how should CF be applied to <see cref="Areas"/>.
    /// </summary>
    /// <remarks>Doesn't seem to work, Excel has no dialogue, nothing found on web and Excel tries
    ///     to repair on row/column values. Avoid if possible.</remarks>
    internal XLPivotCfRuleType Type { get; init; } = XLPivotCfRuleType.None;

    /// <summary>
    /// Areas of pivot table the rule should be applied. The areas are projected to the sheet
    /// <see cref="XLConditionalFormat.Ranges"/> that Excel actually uses to display CF.
    /// </summary>
    internal IReadOnlyList<XLPivotArea> Areas => _area;

    /// <summary>
    /// Conditional format applied to the <see cref="Areas"/>.
    /// </summary>
    /// <remarks>
    /// The <see cref="XLConditionalFormat.Priority"/> of the format is used as a identifier used
    /// to connect pivot CF element and sheet CF element. Pivot CF is ultimately part of sheet CFs
    /// and the priority determines order of CF application (note that CF has
    /// <see cref="XLConditionalFormat.StopIfTrue"/> flag).
    /// </remarks>
    internal XLConditionalFormat Format { get; }

    internal void AddArea(XLPivotArea pivotArea)
    {
        _area.Add(pivotArea);
    }
}
