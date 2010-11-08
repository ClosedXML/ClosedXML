using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public enum XLCellValues { Text, Number, Boolean, DateTime }

    public interface IXLCell: IXLStylized
    {
        Object Value { get; set; }
        IXLAddress Address { get; }
        XLCellValues DataType { get; set; }
        T GetValue<T>();
        String GetString();
        String GetFormattedString();
        Double GetDouble();
        Boolean GetBoolean();
        DateTime GetDateTime();
        void Clear();
        void Delete(XLShiftDeletedCells shiftDeleteCells);
    }
}
