using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal abstract class AbstractPivotFieldReference
    {
        public bool DefaultSubtotal { get; set; }

        internal abstract UInt32Value GetFieldOffset();

        /// <summary>
        ///   <P>Helper function used during saving to calculate the indices of the filtered values</P>
        /// </summary>
        /// <returns>Indices of the filtered values</returns>
        internal abstract IEnumerable<int> Match(XLWorkbook.PivotTableInfo pti, IXLPivotTable pt);
    }
}
