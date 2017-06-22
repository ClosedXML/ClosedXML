using System;
using System.Collections.Generic;
namespace ClosedXML.Excel
{

    public interface IXLTableRange : IXLRange
    {
        IXLTableRow FirstRow(Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRow FirstRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRow FirstRowUsed(Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRow LastRow(Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRow LastRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRow LastRowUsed(Func<IXLTableRow, Boolean> predicate = null);

        new IXLTableRow Row(int row);
        IXLTableRows Rows(Func<IXLTableRow, Boolean> predicate = null);
        new IXLTableRows Rows(int firstRow, int lastRow);
        new IXLTableRows Rows(string rows);
        IXLTableRows RowsUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null);
        IXLTableRows RowsUsed(Func<IXLTableRow, Boolean> predicate = null);

        IXLTable Table { get; }

        new IXLTableRows InsertRowsAbove(int numberOfRows);
        new IXLTableRows InsertRowsBelow(int numberOfRows);
    }
}