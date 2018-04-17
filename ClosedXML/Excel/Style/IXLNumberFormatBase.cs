using System;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormatBase
    {
        Int32 NumberFormatId { get; set; }

        String Format { get; set; }
    }
}
