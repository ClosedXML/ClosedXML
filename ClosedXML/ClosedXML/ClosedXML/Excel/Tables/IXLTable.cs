using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLTableTheme
    {
        TableStyleMedium28,
        TableStyleMedium27,
        TableStyleMedium26,
        TableStyleMedium25,
        TableStyleMedium24,
        TableStyleMedium23,
        TableStyleMedium22,
        TableStyleMedium21,
        TableStyleMedium20,
        TableStyleMedium19,
        TableStyleMedium18,
        TableStyleMedium17,
        TableStyleMedium16,
        TableStyleMedium15,
        TableStyleMedium14,
        TableStyleMedium13,
        TableStyleMedium12,
        TableStyleMedium11,
        TableStyleMedium10,
        TableStyleMedium9,
        TableStyleMedium8,
        TableStyleMedium7,
        TableStyleMedium6,
        TableStyleMedium5,
        TableStyleMedium4,
        TableStyleMedium3,
        TableStyleMedium2,
        TableStyleMedium1,
        TableStyleLight21,
        TableStyleLight20,
        TableStyleLight19,
        TableStyleLight18,
        TableStyleLight17,
        TableStyleLight16,
        TableStyleLight15,
        TableStyleLight14,
        TableStyleLight13,
        TableStyleLight12,
        TableStyleLight11,
        TableStyleLight10,
        TableStyleLight9,
        TableStyleLight8,
        TableStyleLight7,
        TableStyleLight6,
        TableStyleLight5,
        TableStyleLight4,
        TableStyleLight3,
        TableStyleLight2,
        TableStyleLight1,
        TableStyleDark11,
        TableStyleDark10,
        TableStyleDark9,
        TableStyleDark8,
        TableStyleDark7,
        TableStyleDark6,
        TableStyleDark5,
        TableStyleDark4,
        TableStyleDark3,
        TableStyleDark2,
        TableStyleDark1
    }
    public interface IXLTable: IXLRange
    {
        String Name { get; set; }
        Boolean EmphasizeFirstColumn { get; set; }
        Boolean EmphasizeLastColumn { get; set; }
        Boolean ShowRowStripes { get; set; }
        Boolean ShowColumnStripes { get; set; }
        Boolean ShowTotalsRow { get; set; }
        Boolean ShowAutoFilter { get; set; }
        XLTableTheme Theme { get; set; }
        IXLRangeRow HeadersRow();
        IXLRangeRow TotalsRow();
        IXLTableField Field(String fieldName);
        IXLTableField Field(Int32 fieldIndex);

        /// <summary>
        /// Gets the first data row of the table.
        /// </summary>
        new IXLTableRow FirstRow();
        /// <summary>
        /// Gets the first data row of the table that contains a cell with a value.
        /// </summary>
        new IXLTableRow FirstRowUsed();
        /// <summary>
        /// Gets the last data row of the table.
        /// </summary>
        new IXLTableRow LastRow();
        /// <summary>
        /// Gets the last data row of the table that contains a cell with a value.
        /// </summary>
        new IXLTableRow LastRowUsed();
        /// <summary>
        /// Gets the specified row of the table data.
        /// </summary>
        /// <param name="row">The table row.</param>
        new IXLTableRow Row(int row);
        /// <summary>
        /// Gets a collection of all data rows in this table.
        /// </summary>
        new IXLTableRows Rows();
        /// <summary>
        /// Gets a collection of the specified data rows in this table.
        /// </summary>
        /// <param name="firstRow">The first row to return.</param>
        /// <param name="lastRow">The last row to return.</param>
        /// <returns></returns>
        new IXLTableRows Rows(int firstRow, int lastRow);
        /// <summary>
        /// Gets a collection of the specified data rows in this table, separated by commas.
        /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
        /// </summary>
        /// <param name="rows">The rows to return.</param>
        new IXLTableRows Rows(string rows);

        void CopyTo(IXLCell target);
        void CopyTo(IXLRangeBase target);
    }
}
