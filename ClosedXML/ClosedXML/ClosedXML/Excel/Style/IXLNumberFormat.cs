using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormat: IEquatable<IXLNumberFormat>
    {
        Int32 NumberFormatId { get; set; }
        String Format { get; set; }
        IXLStyle SetNumberFormatId(Int32 value);
        IXLStyle SetFormat(String value);
    }
}
