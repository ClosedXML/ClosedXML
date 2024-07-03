using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// A rule describing a subset of pivot table. Used mostly for styling through <see cref="XLPivotFormat"/>.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.3.1.68 PivotArea
/// </remarks>
internal class XLPivotArea
{
    private readonly List<XLPivotReference> _references = new();

    /// <summary>
    /// A subset of field values that are part of the pivot area.
    /// </summary>
    internal IReadOnlyList<XLPivotReference> References => _references;

    /// <summary>
    /// Index of the field that this selection rule refers to.
    /// </summary>
    internal FieldIndex? Field { get; init; }

    /// <summary>
    /// An area of aspect of pivot table that is part of the pivot area.
    /// </summary>
    internal XLPivotAreaType Type { get; init; } = XLPivotAreaType.Normal;

    /// <summary>
    /// Flag indicating whether only the data values (in the data area of the view) for an item
    /// selection are selected and does not include the item labels. Can't be set with together
    /// with <see cref="LabelOnly"/>.
    /// </summary>
    internal bool DataOnly { get; init; } = true;

    /// <summary>
    /// Flag indicating whether only the item labels for an item selection are selected and does
    /// not include the data values(in the data area of the view). Can't be set with together
    /// with <see cref="DataOnly"/>.
    /// </summary>
    internal bool LabelOnly { get; init; } = false;

    /// <summary>
    /// Flag indicating whether the row grand total is included.
    /// </summary>
    internal bool GrandRow { get; init; } = false;

    /// <summary>
    /// Flag indicating whether the column grand total is included.
    /// </summary>
    internal bool GrandCol { get; init; } = false;

    /// <summary>
    /// Flag indicating whether indexes refer to fields or items in the pivot cache and not the
    /// view.
    /// </summary>
    internal bool CacheIndex { get; init; } = false;

    /// <summary>
    /// Flag indicating whether the rule refers to an area that is in outline mode.
    /// </summary>
    internal bool Outline { get; init; } = true;

    /// <summary>
    /// A reference that specifies a subset of the selection area. Points are relative to the top
    /// left of the selection area.
    /// </summary>
    internal XLSheetRange? Offset { get; init; }

    /// <summary>
    /// Flag indicating if collapsed levels/dimensions are considered subtotals.
    /// </summary>
    internal bool CollapsedLevelsAreSubtotals { get; init; } = false;

    /// <summary>
    /// The region of the pivot table to which this rule applies.
    /// </summary>
    internal XLPivotAxis? Axis { get; init; }

    /// <summary>
    /// Position of the field within the axis to which this rule applies.
    /// </summary>
    internal uint? FieldPosition { get; init; }

    internal void AddReference(XLPivotReference reference)
    {
        _references.Add(reference);
    }
}
