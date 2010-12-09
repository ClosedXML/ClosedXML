using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public enum XLCellValues { Text, Number, Boolean, DateTime, TimeSpan }

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
        TimeSpan GetTimeSpan();
        void Clear();
        void Delete(XLShiftDeletedCells shiftDeleteCells);
        String FormulaA1 { get; set; }
        String FormulaR1C1 { get; set; }
        IXLRange AsRange();
    }
}
