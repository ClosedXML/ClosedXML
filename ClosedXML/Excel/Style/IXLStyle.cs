#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLStyle : IEquatable<IXLStyle>
    {
        IXLAlignment Alignment { get; set; }

        IXLBorder Border { get; set; }

        IXLNumberFormat DateFormat { get; }

        IXLFill Fill { get; set; }

        IXLFont Font { get; set; }

        /// <summary>
        /// Should the text values of a cell saved to the file be prefixed by a quote (<c>'</c>) character?
        /// Has no effect if cell values is not a <see cref="XLDataType.Text"/>. Doesn't affect values during runtime,
        /// text values are returned without quote.
        /// </summary>
        Boolean IncludeQuotePrefix { get; set; }

        IXLNumberFormat NumberFormat { get; set; }

        IXLProtection Protection { get; set; }

        IXLStyle SetIncludeQuotePrefix(Boolean includeQuotePrefix = true);
    }
}
