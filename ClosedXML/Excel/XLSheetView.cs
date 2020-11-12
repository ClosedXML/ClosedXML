// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSheetView : IXLSheetView
    {
        private XLAddress _topLeftCellAddress;
        private int _zoomScale;

        public XLSheetView(XLWorksheet worksheet)
        {
            Worksheet = worksheet;
            View = XLSheetViewOptions.Normal;

            ZoomScale = 100;
            ZoomScaleNormal = 100;
            ZoomScalePageLayoutView = 100;
            ZoomScaleSheetLayoutView = 100;
        }

        public XLSheetView(XLWorksheet worksheet, XLSheetView sheetView)
            : this(worksheet)
        {
            this.SplitRow = sheetView.SplitRow;
            this.SplitColumn = sheetView.SplitColumn;
            this.FreezePanes = sheetView.FreezePanes;
            this.TopLeftCellAddress = new XLAddress(this.Worksheet, sheetView.TopLeftCellAddress.RowNumber, sheetView.TopLeftCellAddress.ColumnNumber, sheetView.TopLeftCellAddress.FixedRow, sheetView.TopLeftCellAddress.FixedColumn);
        }

        public Boolean FreezePanes { get; set; }
        public Int32 SplitColumn { get; set; }
        public Int32 SplitRow { get; set; }

        IXLAddress IXLSheetView.TopLeftCellAddress { get => TopLeftCellAddress; set => TopLeftCellAddress = (XLAddress)value; }

        public XLAddress TopLeftCellAddress
        {
            get => _topLeftCellAddress;
            set
            {
                if (value.HasWorksheet && !value.Worksheet.Equals(this.Worksheet))
                    throw new ArgumentException($"The value should be on the same worksheet as the sheet view.");

                _topLeftCellAddress = value;
            }
        }

        public XLSheetViewOptions View { get; set; }

        IXLWorksheet IXLSheetView.Worksheet { get => Worksheet; }
        public XLWorksheet Worksheet { get; internal set; }

        public int ZoomScale
        {
            get { return _zoomScale; }
            set
            {
                _zoomScale = value;
                switch (View)
                {
                    case XLSheetViewOptions.Normal:
                        ZoomScaleNormal = value;
                        break;

                    case XLSheetViewOptions.PageBreakPreview:
                        ZoomScalePageLayoutView = value;
                        break;

                    case XLSheetViewOptions.PageLayout:
                        ZoomScaleSheetLayoutView = value;
                        break;
                }
            }
        }

        public int ZoomScaleNormal { get; set; }

        public int ZoomScalePageLayoutView { get; set; }

        public int ZoomScaleSheetLayoutView { get; set; }

        public void Freeze(Int32 rows, Int32 columns)
        {
            SplitRow = rows;
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeColumns(Int32 columns)
        {
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeRows(Int32 rows)
        {
            SplitRow = rows;
            FreezePanes = true;
        }

        public IXLSheetView SetView(XLSheetViewOptions value)
        {
            View = value;
            return this;
        }
    }
}
