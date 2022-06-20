// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public interface IXLRangeAddress
    {
        /// <summary>
        /// Gets the number of columns in the area covered by the range address.
        /// </summary>
        int ColumnSpan { get; }

        /// <summary>
        /// Gets or sets the first address in the range.
        /// </summary>
        /// <value>
        /// The first address.
        /// </value>
        IXLAddress FirstAddress { get; }

        /// <summary>
        /// Gets or sets a value indicating whether this range is valid.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </value>
        bool IsValid { get; }

        /// <summary>
        /// Gets or sets the last address in the range.
        /// </summary>
        /// <value>
        /// The last address.
        /// </value>
        IXLAddress LastAddress { get; }

        /// <summary>
        /// Gets the number of cells in the area covered by the range address.
        /// </summary>
        int NumberOfCells { get; }

        /// <summary>
        /// Gets the number of rows in the area covered by the range address.
        /// </summary>
        int RowSpan { get; }

        IXLWorksheet Worksheet { get; }

        /// <summary>Allocates the current range address in the internal range repository and returns it</summary>
        IXLRange AsRange();

        bool Contains(IXLAddress address);

        /// <summary>
        /// Returns the intersection of this range address with another range address on the same worksheet.
        /// </summary>
        /// <param name="otherRangeAddress">The other range address.</param>
        /// <returns>The intersection's range address</returns>
        IXLRangeAddress Intersection(IXLRangeAddress otherRangeAddress);

        bool Intersects(IXLRangeAddress otherAddress);

        /// <summary>
        /// Determines whether range address spans the entire column.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire column; otherwise, <c>false</c>.
        /// </returns>
        bool IsEntireColumn();

        /// <summary>
        /// Determines whether range address spans the entire row.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire row; otherwise, <c>false</c>.
        /// </returns>
        bool IsEntireRow();

        /// <summary>
        /// Determines whether the range address spans the entire worksheet.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire sheet; otherwise, <c>false</c>.
        /// </returns>
        bool IsEntireSheet();

        /// <summary>
        /// Returns a range address so that its offset from the target base address is equal to the offset of the current range address to the source base address.
        /// For example, if the current range address is D4:E4, the source base address is A1:C3, then the relative address to the target base address B10:D13 is E14:F14
        /// </summary>
        /// <param name="sourceRangeAddress">The source base range address.</param>
        /// <param name="targetRangeAddress">The target base range address.</param>
        /// <returns>The relative range</returns>
        IXLRangeAddress Relative(IXLRangeAddress sourceRangeAddress, IXLRangeAddress targetRangeAddress);

        string ToString(XLReferenceStyle referenceStyle);

        string ToString(XLReferenceStyle referenceStyle, bool includeSheet);

        string ToStringFixed();

        string ToStringFixed(XLReferenceStyle referenceStyle);

        string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet);

        string ToStringRelative();

        string ToStringRelative(bool includeSheet);
    }
}
