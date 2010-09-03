using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public enum XLCellValues { Text, Number, Boolean, DateTime }

    public interface IXLCell: IXLStylized
    {
        String Value { get; set; }
        IXLAddress Address { get; }
        XLCellValues DataType { get; set; }
    }
}
