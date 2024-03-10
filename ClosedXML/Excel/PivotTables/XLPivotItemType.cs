namespace ClosedXML.Excel;

/// <summary>
/// A categorization of <see cref="XLPivotFieldItem"/> or <see cref="XLPivotFieldAxisItem"/>.
/// </summary>
/// <remarks>
/// 18.18.43 ST_ItemType (PivotItem Type).
/// </remarks>>
internal enum XLPivotItemType
{
    /// <summary>
    /// The pivot item represents an "average" aggregate function.
    /// </summary>
    Avg,

    /// <summary>
    /// The pivot item represents a blank line.
    /// </summary>
    Blank,

    /// <summary>
    /// The pivot item represents custom the "count numbers" aggregate.
    /// </summary>
    Count,

    /// <summary>
    /// The pivot item represents the "count" aggregate function (i.e. number, text and everything
    /// else, except blanks).
    /// </summary>
    CountA,

    /// <summary>
    /// The pivot item represents data.
    /// </summary>
    Data,

    /// <summary>
    /// The pivot item represents the default type for this pivot table, i.e. the "total" aggregate function.
    /// </summary>
    Default,

    /// <summary>
    /// The pivot items represents the grand total line.
    /// </summary>
    Grand,

    /// <summary>
    /// The pivot item represents the "maximum" aggregate function.
    /// </summary>
    Max,

    /// <summary>
    /// The pivot item represents the "minimum" aggregate function.
    /// </summary>
    Min,

    /// <summary>
    /// The pivot item represents the "product" function.
    /// </summary>
    Product,

    /// <summary>
    /// The pivot item represents the "standard deviation" aggregate function.
    /// </summary>
    StdDev,

    /// <summary>
    /// The pivot item represents the "standard deviation population" aggregate function.
    /// </summary>
    StdDevP,

    /// <summary>
    /// The pivot item represents the "sum" aggregate value.
    /// </summary>
    Sum,

    /// <summary>
    /// The pivot item represents the "variance" aggregate value.
    /// </summary>
    Var,

    /// <summary>
    /// The pivot item represents the "variance population" aggregate value.
    /// </summary>
    VarP
}
