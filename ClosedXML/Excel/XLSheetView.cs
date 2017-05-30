using System;

namespace ClosedXML.Excel
{
    internal class XLSheetView : IXLSheetView
    {
        public XLSheetView()
        {
            View = XLSheetViewOptions.Normal;

            ZoomScale = 100;
            ZoomScaleNormal = 100;
            ZoomScalePageLayoutView = 100;
            ZoomScaleSheetLayoutView = 100;
        }

        public XLSheetView(IXLSheetView sheetView)
            : this()
        {
            this.SplitRow = sheetView.SplitRow;
            this.SplitColumn = sheetView.SplitColumn;
            this.FreezePanes = ((XLSheetView)sheetView).FreezePanes;
        }

        public Boolean FreezePanes { get; set; }
        public Int32 SplitColumn { get; set; }
        public Int32 SplitRow { get; set; }
        public XLSheetViewOptions View { get; set; }

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

        private int _zoomScale { get; set; }

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
