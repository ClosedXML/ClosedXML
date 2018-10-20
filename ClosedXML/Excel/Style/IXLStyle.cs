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

        Boolean IncludeQuotePrefix { get; set; }

        IXLNumberFormat NumberFormat { get; set; }

        IXLProtection Protection { get; set; }

        IXLStyle SetIncludeQuotePrefix(Boolean includeQuotePrefix = true);
    }
}
