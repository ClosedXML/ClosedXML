using System;

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

        new IXLTableRow RowAbove();
        new IXLTableRow RowAbove(Int32 step);
        new IXLTableRow RowBelow();
        new IXLTableRow RowBelow(Int32 step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);
    }
}
