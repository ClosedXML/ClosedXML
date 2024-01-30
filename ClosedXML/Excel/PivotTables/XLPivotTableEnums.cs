namespace ClosedXML.Excel
{
    public enum XLFilterAreaOrder { DownThenOver, OverThenDown }

    /// <summary>
    /// Specifies the number of unused items to allow in a <see cref="IXLPivotCache"/>
    /// before discarding unused items.
    /// </summary>
    public enum XLItemsToRetain
    {
        /// <summary>
        /// The threshold is set automatically based on the number of items.
        /// </summary>
        /// <remarks>Default behavior.</remarks>
        Automatic,

        /// <summary>
        /// When even one item is unused.
        /// </summary>
        None,

        /// <summary>
        /// When all shared items of a filed are unused.
        /// </summary>
        Max
    }

    /// <summary>
    /// An enum describing how are values of a <see cref="XLPivotTableField">pivot field</see> are sorted.
    /// </summary>
    /// <remarks>
    /// [ISO-29500] 18.18.28 ST_FieldSortType.
    /// </remarks>
    public enum XLPivotSortType
    {
        /// <summary>
        /// Field values are sorted manually.
        /// </summary>
        Default = 0,

        /// <summary>
        /// Field values are sorted in ascending order.
        /// </summary>
        Ascending = 1,

        /// <summary>
        /// Field values are sorted in descending order.
        /// </summary>
        Descending = 2
    }

    public enum XLPivotSubtotals
    {
        DoNotShow,
        AtTop,
        AtBottom
    }

    public enum XLPivotTableTheme
    {
        None,
        PivotStyleDark1,
        PivotStyleDark10,
        PivotStyleDark11,
        PivotStyleDark12,
        PivotStyleDark13,
        PivotStyleDark14,
        PivotStyleDark15,
        PivotStyleDark16,
        PivotStyleDark17,
        PivotStyleDark18,
        PivotStyleDark19,
        PivotStyleDark2,
        PivotStyleDark20,
        PivotStyleDark21,
        PivotStyleDark22,
        PivotStyleDark23,
        PivotStyleDark24,
        PivotStyleDark25,
        PivotStyleDark26,
        PivotStyleDark27,
        PivotStyleDark28,
        PivotStyleDark3,
        PivotStyleDark4,
        PivotStyleDark5,
        PivotStyleDark6,
        PivotStyleDark7,
        PivotStyleDark8,
        PivotStyleDark9,
        PivotStyleLight1,
        PivotStyleLight10,
        PivotStyleLight11,
        PivotStyleLight12,
        PivotStyleLight13,
        PivotStyleLight14,
        PivotStyleLight15,
        PivotStyleLight16,
        PivotStyleLight17,
        PivotStyleLight18,
        PivotStyleLight19,
        PivotStyleLight2,
        PivotStyleLight20,
        PivotStyleLight21,
        PivotStyleLight22,
        PivotStyleLight23,
        PivotStyleLight24,
        PivotStyleLight25,
        PivotStyleLight26,
        PivotStyleLight27,
        PivotStyleLight28,
        PivotStyleLight3,
        PivotStyleLight4,
        PivotStyleLight5,
        PivotStyleLight6,
        PivotStyleLight7,
        PivotStyleLight8,
        PivotStyleLight9,
        PivotStyleMedium1,
        PivotStyleMedium10,
        PivotStyleMedium11,
        PivotStyleMedium12,
        PivotStyleMedium13,
        PivotStyleMedium14,
        PivotStyleMedium15,
        PivotStyleMedium16,
        PivotStyleMedium17,
        PivotStyleMedium18,
        PivotStyleMedium19,
        PivotStyleMedium2,
        PivotStyleMedium20,
        PivotStyleMedium21,
        PivotStyleMedium22,
        PivotStyleMedium23,
        PivotStyleMedium24,
        PivotStyleMedium25,
        PivotStyleMedium26,
        PivotStyleMedium27,
        PivotStyleMedium28,
        PivotStyleMedium3,
        PivotStyleMedium4,
        PivotStyleMedium5,
        PivotStyleMedium6,
        PivotStyleMedium7,
        PivotStyleMedium8,
        PivotStyleMedium9
    }

    internal enum XLPivotTableSourceType
    {
        /// <summary>
        /// A range in a sheet of the workbook.
        /// </summary>
        Area,

        /// <summary>
        /// Book-scoped named range or a table.
        /// </summary>
        Named
    }
}
