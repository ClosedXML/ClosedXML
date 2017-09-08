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
        IXLAddress FirstAddress { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this range is invalid.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is invalid; otherwise, <c>false</c>.
        /// </value>
        Boolean IsInvalid { get; set; }

        /// <summary>
        /// Gets or sets the last address in the range.
        /// </summary>
        /// <value>
        /// The last address.
        /// </value>
        IXLAddress LastAddress { get; set; }

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
