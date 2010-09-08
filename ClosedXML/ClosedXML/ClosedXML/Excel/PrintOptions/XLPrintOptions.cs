using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLPrintOptions : IXLPrintOptions
    {
        public XLPrintOptions(XLPrintOptions defaultPrintOptions)
        {
            if (defaultPrintOptions != null)
            {
                this.CenterHorizontally = defaultPrintOptions.CenterHorizontally;
                this.CenterVertically = defaultPrintOptions.CenterVertically;
                this.FirstPageNumber = defaultPrintOptions.FirstPageNumber;
                this.HorizontalDpi = defaultPrintOptions.HorizontalDpi;
                this.PageOrientation = defaultPrintOptions.PageOrientation;
                this.VerticalDpi = defaultPrintOptions.VerticalDpi;
                if (defaultPrintOptions.PrintArea != null)
                {
                    this.PrintArea = new XLRange(
                        new XLRangeParameters(
                            defaultPrintOptions.PrintArea.Internals.FirstCellAddress,
                            defaultPrintOptions.PrintArea.Internals.LastCellAddress,
                            defaultPrintOptions.PrintArea.Internals.Worksheet,
                            defaultPrintOptions.PrintArea.Style)
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

                if (defaultPrintOptions.HeaderFooters != null)
                {
                    //HeaderFooters = new XLHeaderFooter();
                }
            }
        }
        public IXLRange PrintArea { get; set; }
        public XLPageOrientation PageOrientation { get; set; }
        public XLPaperSize PaperSize { get; set; }
        public Int32 HorizontalDpi { get; set; }
        public Int32 VerticalDpi { get; set; }
        public Int32 FirstPageNumber { get; set; }
        public Boolean CenterHorizontally { get; set; }
        public Boolean CenterVertically { get; set; }

        public XLMargins Margins { get; set; }
        public IXLHeaderFooter HeaderFooters { get; private set; }

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
                pagesTall = 0;
                pagesWide = 0;
            }
        }

        public void AdjustTo(Int32 pctOfNormalSize)
        {
            Scale = pctOfNormalSize;
            pagesWide = 0;
            pagesTall = 0;
        }
        public void FitTo(Int32 pagesWide, Int32 pagesTall)
        {
            this.pagesWide = pagesWide;
            this.pagesTall = pagesTall;
            Scale = 0;
        }
    }
}
