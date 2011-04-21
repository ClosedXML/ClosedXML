using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLTableRow: IXLRangeRow
    {
        IXLCell Field(Int32 index);
        IXLCell Field(String name);

        new IXLTableRow Sort();
        new IXLTableRow Sort(Boolean matchCase);
        new IXLTableRow Sort(XLSortOrder sortOrder);
        new IXLTableRow Sort(XLSortOrder sortOrder, Boolean matchCase);

        new IXLTableRow Replace(String oldValue, String newValue);
        new IXLTableRow Replace(String oldValue, String newValue, XLSearchContents searchContents);
        new IXLTableRow Replace(String oldValue, String newValue, XLSearchContents searchContents, Boolean useRegularExpressions);
    }
}
