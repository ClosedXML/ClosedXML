using System;
using System.Collections.Generic;
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
        TableStyleDark1,
        None
    }

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