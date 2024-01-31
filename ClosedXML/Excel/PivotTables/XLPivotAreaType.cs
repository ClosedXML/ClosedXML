namespace ClosedXML.Excel
{
    /// <summary>
    /// An area of aspect of pivot table that is part of the <see cref="XLPivotArea.Type"/>.
    /// </summary>
    /// <remarks>
    /// [ISO-29500] 18.18.58 ST_PivotAreaType
    /// </remarks>
    internal enum XLPivotAreaType
    {
        None = 0,
        Normal = 1,
        Data = 2,
        All = 3,
        Origin = 4,
        Button = 5,

        // Top right has been removed between ISO-29500:2006 and ISO-29500:2016.
        TopRight = 6,
        TopEnd = 7
    }
}
