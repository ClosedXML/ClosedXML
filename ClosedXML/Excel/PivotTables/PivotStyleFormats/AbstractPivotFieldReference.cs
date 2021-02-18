using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal abstract class AbstractPivotFieldReference
    {
        internal XLPivotStyleFormatSubTotalFilter SubTotalFilter { get; set; } = XLPivotStyleFormatSubTotalFilter.None;

        internal abstract UInt32Value GetFieldOffset();

        /// <summary>
        ///   <P>Helper function used during saving to calculate the indices of the filtered values</P>
        /// </summary>
        /// <returns>Indices of the filtered values</returns>
        internal abstract IEnumerable<Int32> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt);
    }
}
