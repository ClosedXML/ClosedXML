namespace ClosedXML.Excel
{
    public interface IXLTableRow : IXLRangeRow
    {
        IXLCell Field(int index);

        IXLCell Field(string name);

        new IXLTableRow Sort();

        new IXLTableRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        new IXLTableRow RowAbove();

        new IXLTableRow RowAbove(int step);

        new IXLTableRow RowBelow();

        new IXLTableRow RowBelow(int step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        new IXLTableRows InsertRowsAbove(int numberOfRows);

        new IXLTableRows InsertRowsBelow(int numberOfRows);
    }
}
