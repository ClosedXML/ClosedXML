using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// A representation of a single row/column axis values in a <see cref="XLPivotTableAxis"/>.
/// </summary>
/// <remarks>
/// Represents 18.10.1.44 i (Row Items) and 18.10.1.96 x (Member Property Index).
/// </remarks>
internal class XLPivotFieldAxisItem
{
    public XLPivotFieldAxisItem(XLPivotItemType itemType, int dataItem, IEnumerable<int> fieldItems)
    {
        ItemType = itemType;
        DataItem = dataItem;
        FieldItem = fieldItems.ToList();
    }

    /// <summary>
    /// Each item is an index to field items of corresponding field from
    /// <see cref="XLPivotTableAxis.Fields"/>. Value <c>1048832</c> specifies that no item appears
    /// at the position. 
    /// </summary>
    internal List<int> FieldItem { get; }

    /// <summary>
    /// Type of item.
    /// </summary>
    internal XLPivotItemType ItemType { get; }

    /// <summary>
    /// If this item (row/column) contains 'data' field, this contains an index into the <see cref="XLPivotTable.DataFields"/>
    /// that should be used as a value. The value for 'data' field in the <see cref="FieldItem"/> is ignored, but Excel fills
    /// same number as this index.
    /// </summary>
    internal int DataItem { get; }
}
