// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    internal class XLPivotTableStyleFormats : IXLPivotTableStyleFormats
    {
        private IXLPivotStyleFormats columnGrandTotalFormats;
        private IXLPivotStyleFormats rowGrandTotalFormats;

        #region IXLPivotTableStyleFormats members

        public IXLPivotStyleFormats ColumnGrandTotalFormats
        {
            get { return columnGrandTotalFormats ?? (columnGrandTotalFormats = new XLPivotStyleFormats()); }
            set { columnGrandTotalFormats = value; }
        }

        public IXLPivotStyleFormats RowGrandTotalFormats
        {
            get { return rowGrandTotalFormats ?? (rowGrandTotalFormats = new XLPivotStyleFormats()); }
            set { rowGrandTotalFormats = value; }
        }

        #endregion IXLPivotTableStyleFormats members
    }
}
