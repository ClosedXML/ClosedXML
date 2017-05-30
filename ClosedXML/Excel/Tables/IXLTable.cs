using System;
using System.Collections.Generic;
namespace ClosedXML.Excel
{
    public interface IXLTable : IXLRange
    {
        string Name { get; set; }
        Boolean EmphasizeFirstColumn { get; set; }
        Boolean EmphasizeLastColumn { get; set; }
        Boolean ShowRowStripes { get; set; }
        Boolean ShowColumnStripes { get; set; }
        Boolean ShowTotalsRow { get; set; }
        Boolean ShowAutoFilter { get; set; }
        XLTableTheme Theme { get; set; }
        IXLRangeRow HeadersRow();
        IXLRangeRow TotalsRow();
        IXLTableField Field(string fieldName);
        IXLTableField Field(int fieldIndex);
        IEnumerable<IXLTableField> Fields { get; }

       

        IXLTable SetEmphasizeFirstColumn(); IXLTable SetEmphasizeFirstColumn(Boolean value);
        IXLTable SetEmphasizeLastColumn(); IXLTable SetEmphasizeLastColumn(Boolean value);
        IXLTable SetShowRowStripes(); IXLTable SetShowRowStripes(Boolean value);
        IXLTable SetShowColumnStripes(); IXLTable SetShowColumnStripes(Boolean value);
        IXLTable SetShowTotalsRow(); IXLTable SetShowTotalsRow(Boolean value);
        IXLTable SetShowAutoFilter(); IXLTable SetShowAutoFilter(Boolean value);

        /// <summary>
        /// Clears the contents of this table.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        IXLBaseAutoFilter AutoFilter { get; }

        new IXLBaseAutoFilter SetAutoFilter();

        Boolean ShowHeaderRow { get; set; }
        IXLTable SetShowHeaderRow(); IXLTable SetShowHeaderRow(Boolean value);

        IXLTableRange DataRange { get; }
    }
}