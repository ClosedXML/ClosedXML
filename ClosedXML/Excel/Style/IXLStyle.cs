using System;

namespace ClosedXML.Excel
{
    public interface IXLStyle : IEquatable<IXLStyle>
    {
        IXLAlignment Alignment { get; set; }

        IXLBorder Border { get; set; }

        IXLFill Fill { get; set; }

        IXLFont Font { get; set; }

        IXLNumberFormat NumberFormat { get; set; }

        IXLNumberFormat DateFormat { get; }

        IXLProtection Protection { get; set; }
    }
}
