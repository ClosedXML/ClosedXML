using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Drawings
{
    public interface IXLMarker
    {
        Int32 ColumnId { get; set; }
        Int32 RowId { get; set; }
        Double ColumnOffset { get; set; }
        Double RowOffset { get; set; }

        /// <summary>
        /// Get the zero-based column number.
        /// </summary>
        Int32 GetZeroBasedColumn();

        /// <summary>
        /// Get the zero-based row number.
        /// </summary>
        Int32 GetZeroBasedRow();
    }
}
