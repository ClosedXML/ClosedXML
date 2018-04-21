using System;

namespace ClosedXML.Excel
{
    public interface IXLRangeAddress
    {
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
        Boolean IsValid { get; }

        /// <summary>
        /// Gets or sets the last address in the range.
        /// </summary>
        /// <value>
        /// The last address.
        /// </value>
        IXLAddress LastAddress { get; }

        IXLWorksheet Worksheet { get; }

        String ToString(XLReferenceStyle referenceStyle);

        String ToString(XLReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringFixed();

        String ToStringFixed(XLReferenceStyle referenceStyle);

        String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet);

        String ToStringRelative();

        String ToStringRelative(Boolean includeSheet);

        bool Intersects(IXLRangeAddress otherAddress);

        bool Contains(IXLAddress address);
    }
}
