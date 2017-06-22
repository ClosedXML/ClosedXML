using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLCFIconSetOperator {GreaterThan, EqualOrGreaterThan}
    public interface IXLCFIconSet
    {
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, String value, XLCFContentType type);
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, Double value, XLCFContentType type);
    }
}
