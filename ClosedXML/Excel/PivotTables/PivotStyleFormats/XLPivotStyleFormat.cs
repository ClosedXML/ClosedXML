// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPivotStyleFormat : IXLPivotStyleFormat
    {
        public XLPivotStyleFormat(IXLPivotField field = null, IXLStyle style = null)
        {
            PivotField = field;
            Style = style ?? XLStyle.Default;
        }

        #region IXLPivotStyleFormat members

        public XLPivotStyleFormatElement AppliesTo { get; set; } = XLPivotStyleFormatElement.Data;
        public IXLPivotField PivotField { get; set; }
        public IXLStyle Style { get; set; }

        #endregion IXLPivotStyleFormat members

        internal IList<AbstractPivotFieldReference> FieldReferences { get; } = new List<AbstractPivotFieldReference>();
        internal XLPivotStyleFormatRule Rule { get; } = new XLPivotStyleFormatRule();
    }
}
