using System;

namespace ClosedXML.Excel
{
    internal class XLSheetView: IXLSheetView
    {
        public XLSheetView() {
            View = XLSheetViewOptions.Normal;
        }
        public XLSheetView(IXLSheetView sheetView):this()
        {
            this.SplitRow = sheetView.SplitRow;
            this.SplitColumn = sheetView.SplitColumn;
            this.FreezePanes = ((XLSheetView)sheetView).FreezePanes;
        }

        public Int32 SplitRow { get; set; }
        public Int32 SplitColumn { get; set; }
        public Boolean FreezePanes { get; set; }
        public void FreezeRows(Int32 rows)
        {
            SplitRow = rows;
            FreezePanes = true;
        }
        public void FreezeColumns(Int32 columns)
        {
            SplitColumn = columns;
            FreezePanes = true;
        }
        public void Freeze(Int32 rows, Int32 columns)
        {
            SplitRow = rows;
            SplitColumn = columns;
            FreezePanes = true;
        }


        public XLSheetViewOptions View { get; set; }

        public IXLSheetView SetView(XLSheetViewOptions value)
        {
            View = value;
            return this;
        }
    }
}
