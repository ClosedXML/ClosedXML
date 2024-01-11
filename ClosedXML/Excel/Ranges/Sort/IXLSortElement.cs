using System;

namespace ClosedXML.Excel
{
    public enum XLSortOrder { Ascending, Descending }
    public enum XLSortOrientation { TopToBottom, LeftToRight }
    public interface IXLSortElement
    {
        /// <summary>
        /// Column or row number whose values will be used for sorting.
        /// </summary>
        Int32 ElementNumber { get; }

        /// <summary>
        /// Sorting order.
        /// </summary>
        XLSortOrder SortOrder { get; }

        /// <summary>
        /// When <c>true</c> (recommended, matches Excel behavior), blank cell values are always
        /// sorted at the end regardless of sorting order. When <c>false</c>, blank values are
        /// considered empty strings and are sorted among other cell values with a type
        /// <see cref="XLDataType.Text"/>.
        /// </summary>
        Boolean IgnoreBlanks { get; }

        /// <summary>
        /// When cell value is a <see cref="XLDataType.Text"/>, should sorting be case insensitive
        /// (<c>false</c>, Excel default behavior) or case sensitive (<c>true</c>). Doesn't affect
        /// other cell value types.
        /// </summary>
        Boolean MatchCase { get; }
    }
}
