using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLPageOptions : IXLPageSetup
    {
        public XLPageOptions(XLPageOptions defaultPrintOptions)
        {
            this.PrintAreas = new List<IXLRange>();
            this.RowTitles = new List<IXLRow>();
            this.ColumnTitles = new List<IXLColumn>();
            if (defaultPrintOptions != null)
            {
                this.CenterHorizontally = defaultPrintOptions.CenterHorizontally;
                this.CenterVertically = defaultPrintOptions.CenterVertically;
                this.FirstPageNumber = defaultPrintOptions.FirstPageNumber;
                this.HorizontalDpi = defaultPrintOptions.HorizontalDpi;
                this.PageOrientation = defaultPrintOptions.PageOrientation;
                this.VerticalDpi = defaultPrintOptions.VerticalDpi;
                foreach (var printArea in defaultPrintOptions.PrintAreas)
                {
                    this.PrintAreas.Add(
                        new XLRange(
                        new XLRangeParameters(
                            printArea.Internals.FirstCellAddress,
                            printArea.Internals.LastCellAddress,
                            printArea.Internals.Worksheet,
                            printArea.Style)
                        )
                        );
                }

                foreach (var rowTitle in defaultPrintOptions.RowTitles)
                {
                    this.RowTitles.Add(
                        new XLRow(rowTitle.RowNumber, new XLRowParameters(rowTitle.Internals.Worksheet, rowTitle.Style))
                        );
                }
                foreach (var columnTitle in defaultPrintOptions.ColumnTitles)
                {
                    this.ColumnTitles.Add(
                        new XLColumn(columnTitle.ColumnNumber, new XLColumnParameters(columnTitle.Internals.Worksheet, columnTitle.Style))
                        );
                }
                this.PaperSize = defaultPrintOptions.PaperSize;
                this.pagesTall = defaultPrintOptions.pagesTall;
                this.pagesWide = defaultPrintOptions.pagesWide;
                this.scale = defaultPrintOptions.scale;

                if (defaultPrintOptions.Margins != null)
                {
                    this.Margins = new XLMargins()
                            {
                                Top = defaultPrintOptions.Margins.Top,
                                Bottom = defaultPrintOptions.Margins.Bottom,
                                Left = defaultPrintOptions.Margins.Left,
                                Right = defaultPrintOptions.Margins.Right,
                                Header = defaultPrintOptions.Margins.Header,
                                Footer = defaultPrintOptions.Margins.Footer
                            };
                }
                this.AlignHFWithMargins = defaultPrintOptions.AlignHFWithMargins;
                this.ScaleHFWithDocument = defaultPrintOptions.ScaleHFWithDocument;
                this.ShowGridlines = defaultPrintOptions.ShowGridlines;
                this.ShowRowAndColumnHeadings = defaultPrintOptions.ShowRowAndColumnHeadings;
                this.BlackAndWhite = defaultPrintOptions.BlackAndWhite;
                this.DraftQuality = defaultPrintOptions.DraftQuality;
                this.PageOrder = defaultPrintOptions.PageOrder;

                this.ColumnBreaks = new List<IXLColumn>();
                this.RowBreaks = new List<IXLRow>();
                this.PrintErrorValue = defaultPrintOptions.PrintErrorValue;
            }
            Header = new XLHeaderFooter();
            Footer = new XLHeaderFooter();
        }
        public List<IXLRange> PrintAreas { get; set; }
        public List<IXLRow> RowTitles { get; set; }
        public List<IXLColumn> ColumnTitles { get; set; }
        public void SetRowTitles(List<IXLRow> rowTitles)
        {
            RowTitles.Clear();
            RowTitles.AddRange(rowTitles);
        }
        public void SetColumnTitles(List<IXLColumn> columnTitles)
        {
            ColumnTitles.Clear();
            ColumnTitles.AddRange(columnTitles);
        }
        public XLPageOrientation PageOrientation { get; set; }
        public XLPaperSize PaperSize { get; set; }
        public Int32 HorizontalDpi { get; set; }
        public Int32 VerticalDpi { get; set; }
        public Int32 FirstPageNumber { get; set; }
        public Boolean CenterHorizontally { get; set; }
        public Boolean CenterVertically { get; set; }
        public XLPrintErrorValues PrintErrorValue { get; set; }
        public XLMargins Margins { get; set; }

        private Int32 pagesWide;
        public Int32 PagesWide 
        {
            get
            {
                return pagesWide;
            }
            set
            {
                pagesWide = value;
                if (pagesWide >0)
                    scale = 0;
            }
        }
        
        private Int32 pagesTall;
        public Int32 PagesTall
        {
            get
            {
                return pagesTall;
            }
            set
            {
                pagesTall = value;
                if (pagesTall >0)
                    scale = 0;
            }
        }

        private Int32 scale;
        public Int32 Scale
        {
            get
            {
                return scale;
            }
            set
            {
                scale = value;
                if (scale > 0)
                {
                    pagesTall = 0;
                    pagesWide = 0;
                }
            }
        }

        public void AdjustTo(Int32 pctOfNormalSize)
        {
            Scale = pctOfNormalSize;
            pagesWide = 0;
            pagesTall = 0;
        }
        public void FitToPages(Int32 pagesWide, Int32 pagesTall)
        {
            this.pagesWide = pagesWide;
            this.pagesTall = pagesTall;
            scale = 0;
        }


        public IXLHeaderFooter Header { get; private set; }
        public IXLHeaderFooter Footer { get; private set; }

        public Boolean ScaleHFWithDocument { get; set; }
        public Boolean AlignHFWithMargins { get; set; }

        public Boolean ShowGridlines { get; set; }
        public Boolean ShowRowAndColumnHeadings { get; set; }
        public Boolean BlackAndWhite { get; set; }
        public Boolean DraftQuality { get; set; }

        public XLPageOrderValues PageOrder { get; set; }
        public XLShowCommentsValues ShowComments { get; set; }

        public List<IXLRow> RowBreaks { get; private set; }
        public List<IXLColumn> ColumnBreaks { get; private set; }

        public void AddPageBreak(IXLRow row)
        {
            RowBreaks.Add(row);
        }
        public void AddPageBreak(IXLColumn column)
        {
            ColumnBreaks.Add(column);
        }

        //public void SetPageBreak(IXLRange range, XLPageBreakLocations breakLocation)
        //{
        //    switch (breakLocation)
        //    {
        //        case XLPageBreakLocations.AboveRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowNumber)); break;
        //        case XLPageBreakLocations.BelowRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowCount())); break;
        //        case XLPageBreakLocations.LeftOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnNumber)); break;
        //        case XLPageBreakLocations.RightOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnCount())); break;
        //        default: throw new NotImplementedException();
        //    }
        //}
    }
}
