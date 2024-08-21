using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// Represents a set of selected fields and selected items within those fields. It's used to select
/// an area for <see cref="XLPivotArea"/>.
/// </summary>
internal class XLPivotReference
{
    private readonly List<uint> _fieldItems = new();

    /// <summary>
    /// <para>
    /// If <see cref="XLPivotArea.CacheIndex"/> is <c>false</c>, then it is index into pivot fields
    /// items of pivot field <see cref="Field"/> (unless <see cref="ByPosition"/> is <c>true</c>).
    /// </para>
    /// <para>
    /// If <see cref="XLPivotArea.CacheIndex"/> is <c>true</c>, then it is index into shared items
    /// of a cached field with index <see cref="Field"/> (unless <see cref="ByPosition"/> is
    /// <c>true</c>).
    /// </para>
    /// </summary>
    internal List<uint> FieldItems => _fieldItems;

    /// <summary>
    /// Specifies the index of the field to which this filter refers. A value of -2/4294967294
    /// indicates the 'data' field. It can represent pivot field or cache field, depending on
    /// <see cref="XLPivotArea.CacheIndex"/>.
    /// </summary>
    internal uint? Field { get; init; }

    /// <summary>
    /// Flag indicating whether this field has selection. This attribute is used when the
    /// pivot table is in outline view. It is also used when both header and data
    /// cells have selection.
    /// </summary>
    internal bool Selected { get; init; } = true;

    /// <summary>
    /// Flag indicating whether the item in <see cref="FieldItems"/> is referred to by position rather
    /// than item index.
    /// </summary>
    internal bool ByPosition { get; init; } = false;

    /// <summary>
    /// Flag indicating whether the item is referred to by a relative reference rather than an
    /// absolute reference. This attribute is used if posRef is set to true.
    /// </summary>
    internal bool Relative { get; init; } = false;

    internal HashSet<XLSubtotalFunction> Subtotals { get; init; }= new();

    internal void AddFieldItem(uint fieldItem)
    {
        // TODO: Check value by area.CacheIndex and ByPosition
        _fieldItems.Add(fieldItem);
    }
}
