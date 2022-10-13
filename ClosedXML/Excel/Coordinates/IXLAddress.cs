using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Reference to a single cell in a workbook. Reference can be absolute, relative or mixed.
    /// Reference can be with or without a worksheet.
    /// </summary>
    public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>
    {
        String ColumnLetter { get; }
        Int32 ColumnNumber { get; }
        Boolean FixedColumn { get; }
        Boolean FixedRow { get; }
        Int32 RowNumber { get; }
        String UniqueId { get; }

        /// <summary>
        /// Worksheet of the reference. Value is null for address without a worksheet.
        /// </summary>
        IXLWorksheet Worksheet { get; }

        String ToString(XLReferenceStyle referenceStyle);

        String ToString(XLReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringFixed();

        String ToStringFixed(XLReferenceStyle referenceStyle);

        String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringRelative();

        String ToStringRelative(Boolean includeSheet);
    }
}
