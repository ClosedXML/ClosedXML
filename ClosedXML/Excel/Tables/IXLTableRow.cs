using System;

namespace ClosedXML.Excel
{
    public interface IXLTableRow : IXLRangeRow
    {
        IXLCell Field(Int32 index);

        IXLCell Field(String name);

        new IXLTableRow Sort();

        new IXLTableRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        new IXLTableRow RowAbove();

        new IXLTableRow RowAbove(Int32 step);

        new IXLTableRow RowBelow();

        new IXLTableRow RowBelow(Int32 step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        new IXLTableRows InsertRowsAbove(int numberOfRows);

        new IXLTableRows InsertRowsBelow(int numberOfRows);
    }
}
