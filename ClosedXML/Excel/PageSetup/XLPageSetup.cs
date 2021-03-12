using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPageSetup : IXLPageSetup
    {
        public XLPageSetup(XLPageSetup defaultPageOptions, XLWorksheet worksheet)
        {

            if (defaultPageOptions != null)
            {
                PrintAreas = new XLPrintAreas(defaultPageOptions.PrintAreas as XLPrintAreas, worksheet);
                CenterHorizontally = defaultPageOptions.CenterHorizontally;
                CenterVertically = defaultPageOptions.CenterVertically;
                FirstPageNumber = defaultPageOptions.FirstPageNumber;
                HorizontalDpi = defaultPageOptions.HorizontalDpi;
                PageOrientation = defaultPageOptions.PageOrientation;
                VerticalDpi = defaultPageOptions.VerticalDpi;
                FirstRowToRepeatAtTop = defaultPageOptions.FirstRowToRepeatAtTop;
                LastRowToRepeatAtTop = defaultPageOptions.LastRowToRepeatAtTop;
                FirstColumnToRepeatAtLeft = defaultPageOptions.FirstColumnToRepeatAtLeft;
                LastColumnToRepeatAtLeft = defaultPageOptions.LastColumnToRepeatAtLeft;
                ShowComments = defaultPageOptions.ShowComments;


                PaperSize = defaultPageOptions.PaperSize;
                _pagesTall = defaultPageOptions.PagesTall;
                _pagesWide = defaultPageOptions.PagesWide;
                _scale = defaultPageOptions.Scale;


                if (defaultPageOptions.Margins != null)
                {
                    Margins = new XLMargins
                                  {
                                Top = defaultPageOptions.Margins.Top,
                                Bottom = defaultPageOptions.Margins.Bottom,
                                Left = defaultPageOptions.Margins.Left,
                                Right = defaultPageOptions.Margins.Right,
                                Header = defaultPageOptions.Margins.Header,
                                Footer = defaultPageOptions.Margins.Footer
                            };
                }
                AlignHFWithMargins = defaultPageOptions.AlignHFWithMargins;
                ScaleHFWithDocument = defaultPageOptions.ScaleHFWithDocument;
                ShowGridlines = defaultPageOptions.ShowGridlines;
                ShowRowAndColumnHeadings = defaultPageOptions.ShowRowAndColumnHeadings;
                BlackAndWhite = defaultPageOptions.BlackAndWhite;
                DraftQuality = defaultPageOptions.DraftQuality;
                PageOrder = defaultPageOptions.PageOrder;

                ColumnBreaks = defaultPageOptions.ColumnBreaks.ToList();
                RowBreaks = defaultPageOptions.RowBreaks.ToList();
                Header = new XLHeaderFooter(defaultPageOptions.Header as XLHeaderFooter, worksheet);
                Footer = new XLHeaderFooter(defaultPageOptions.Footer as XLHeaderFooter, worksheet);
                PrintErrorValue = defaultPageOptions.PrintErrorValue;
            }
            else
            {
                PrintAreas = new XLPrintAreas(worksheet);
                Header = new XLHeaderFooter(worksheet);
                Footer = new XLHeaderFooter(worksheet);
                ColumnBreaks = new List<Int32>();
                RowBreaks = new List<Int32>();
            }
        }
        public IXLPrintAreas PrintAreas { get; private set; }


        public Int32 FirstRowToRepeatAtTop { get; private set; }
        public Int32 LastRowToRepeatAtTop { get; private set; }
        public void SetRowsToRepeatAtTop(String range)
        {
            var arrRange = range.Replace("$", "").Split(':');
            SetRowsToRepeatAtTop(Int32.Parse(arrRange[0]), Int32.Parse(arrRange[1]));
        }
        public void SetRowsToRepeatAtTop(Int32 firstRowToRepeatAtTop, Int32 lastRowToRepeatAtTop)
        {
            if (firstRowToRepeatAtTop <= 0) throw new ArgumentOutOfRangeException("The first row has to be greater than zero.");
            if (firstRowToRepeatAtTop > lastRowToRepeatAtTop) throw new ArgumentOutOfRangeException("The first row has to be less than the second row.");

            FirstRowToRepeatAtTop = firstRowToRepeatAtTop;
            LastRowToRepeatAtTop = lastRowToRepeatAtTop;
        }
        public Int32 FirstColumnToRepeatAtLeft { get; private set; }
        public Int32 LastColumnToRepeatAtLeft { get; private set; }
        public void SetColumnsToRepeatAtLeft(String range)
        {
            var arrRange = range.Replace("$", "").Split(':');
            if (Int32.TryParse(arrRange[0], out int iTest))
                SetColumnsToRepeatAtLeft(Int32.Parse(arrRange[0]), Int32.Parse(arrRange[1]));
            else
                SetColumnsToRepeatAtLeft(arrRange[0], arrRange[1]);
        }
        public void SetColumnsToRepeatAtLeft(String firstColumnToRepeatAtLeft, String lastColumnToRepeatAtLeft)
        {
            SetColumnsToRepeatAtLeft(XLHelper.GetColumnNumberFromLetter(firstColumnToRepeatAtLeft), XLHelper.GetColumnNumberFromLetter(lastColumnToRepeatAtLeft));
        }
        public void SetColumnsToRepeatAtLeft(Int32 firstColumnToRepeatAtLeft, Int32 lastColumnToRepeatAtLeft)
        {
            if (firstColumnToRepeatAtLeft <= 0) throw new ArgumentOutOfRangeException("The first column has to be greater than zero.");
            if (firstColumnToRepeatAtLeft > lastColumnToRepeatAtLeft) throw new ArgumentOutOfRangeException("The first column has to be less than the second column.");

            FirstColumnToRepeatAtLeft = firstColumnToRepeatAtLeft;
            LastColumnToRepeatAtLeft = lastColumnToRepeatAtLeft;
        }

        public XLPageOrientation PageOrientation { get; set; }
        public XLPaperSize PaperSize { get; set; }
        public Int32 HorizontalDpi { get; set; }
        public Int32 VerticalDpi { get; set; }
        public UInt32? FirstPageNumber { get; set; }
        public Boolean CenterHorizontally { get; set; }
        public Boolean CenterVertically { get; set; }
        public XLPrintErrorValues PrintErrorValue { get; set; }
        public IXLMargins Margins { get; set; }

        private Int32 _pagesWide;
        public Int32 PagesWide
        {
            get
            {
                return _pagesWide;
            }
            set
            {
                _pagesWide = value;
                if (_pagesWide >0)
                    _scale = 0;
            }
        }

        private Int32 _pagesTall;
        public Int32 PagesTall
        {
            get
            {
                return _pagesTall;
            }
            set
            {
                _pagesTall = value;
                if (_pagesTall >0)
                    _scale = 0;
            }
        }

        private Int32 _scale;
        public Int32 Scale
        {
            get
            {
                return _scale;
            }
            set
            {
                _scale = value;
                if (_scale <= 0) return;
                _pagesTall = 0;
                _pagesWide = 0;
            }
        }

        public void AdjustTo(Int32 percentageOfNormalSize)
        {
            Scale = percentageOfNormalSize;
            _pagesWide = 0;
            _pagesTall = 0;
        }
        public void FitToPages(Int32 pagesWide, Int32 pagesTall)
        {
            _pagesWide = pagesWide;
            this._pagesTall = pagesTall;
            _scale = 0;
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

        public List<Int32> RowBreaks { get; private set; }
        public List<Int32> ColumnBreaks { get; private set; }
        public void AddHorizontalPageBreak(Int32 row)
        {
            if (!RowBreaks.Contains(row))
                RowBreaks.Add(row);
            RowBreaks.Sort();
        }
        public void AddVerticalPageBreak(Int32 column)
        {
            if (!ColumnBreaks.Contains(column))
                ColumnBreaks.Add(column);
            ColumnBreaks.Sort();
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

        public IXLPageSetup SetPageOrientation(XLPageOrientation value) { PageOrientation = value; return this; }
        public IXLPageSetup SetPagesWide(Int32 value) { PagesWide = value; return this; }
        public IXLPageSetup SetPagesTall(Int32 value) { PagesTall = value; return this; }
        public IXLPageSetup SetScale(Int32 value) { Scale = value; return this; }
        public IXLPageSetup SetHorizontalDpi(Int32 value) { HorizontalDpi = value; return this; }
        public IXLPageSetup SetVerticalDpi(Int32 value) { VerticalDpi = value; return this; }
        public IXLPageSetup SetFirstPageNumber(UInt32? value) { FirstPageNumber = value; return this; }
        public IXLPageSetup SetCenterHorizontally() { CenterHorizontally = true; return this; }	public IXLPageSetup SetCenterHorizontally(Boolean value) { CenterHorizontally = value; return this; }
        public IXLPageSetup SetCenterVertically() { CenterVertically = true; return this; }	public IXLPageSetup SetCenterVertically(Boolean value) { CenterVertically = value; return this; }
        public IXLPageSetup SetPaperSize(XLPaperSize value) { PaperSize = value; return this; }
        public IXLPageSetup SetScaleHFWithDocument() { ScaleHFWithDocument = true; return this; }	public IXLPageSetup SetScaleHFWithDocument(Boolean value) { ScaleHFWithDocument = value; return this; }
        public IXLPageSetup SetAlignHFWithMargins() { AlignHFWithMargins = true; return this; }	public IXLPageSetup SetAlignHFWithMargins(Boolean value) { AlignHFWithMargins = value; return this; }
        public IXLPageSetup SetShowGridlines() { ShowGridlines = true; return this; }	public IXLPageSetup SetShowGridlines(Boolean value) { ShowGridlines = value; return this; }
        public IXLPageSetup SetShowRowAndColumnHeadings() { ShowRowAndColumnHeadings = true; return this; }	public IXLPageSetup SetShowRowAndColumnHeadings(Boolean value) { ShowRowAndColumnHeadings = value; return this; }
        public IXLPageSetup SetBlackAndWhite() { BlackAndWhite = true; return this; }	public IXLPageSetup SetBlackAndWhite(Boolean value) { BlackAndWhite = value; return this; }
        public IXLPageSetup SetDraftQuality() { DraftQuality = true; return this; }	public IXLPageSetup SetDraftQuality(Boolean value) { DraftQuality = value; return this; }
        public IXLPageSetup SetPageOrder(XLPageOrderValues value) { PageOrder = value; return this; }
        public IXLPageSetup SetShowComments(XLShowCommentsValues value) { ShowComments = value; return this; }
        public IXLPageSetup SetPrintErrorValue(XLPrintErrorValues value) { PrintErrorValue = value; return this; }

        public Boolean DifferentFirstPageOnHF { get; set; }
        public IXLPageSetup SetDifferentFirstPageOnHF()
        {
            return SetDifferentFirstPageOnHF(true);
        }
        public IXLPageSetup SetDifferentFirstPageOnHF(Boolean value)
        {
            DifferentFirstPageOnHF = value;
            return this;
        }
        public Boolean DifferentOddEvenPagesOnHF { get; set; }
        public IXLPageSetup SetDifferentOddEvenPagesOnHF()
        {
            return SetDifferentOddEvenPagesOnHF(true);
        }
        public IXLPageSetup SetDifferentOddEvenPagesOnHF(Boolean value)
        {
            DifferentOddEvenPagesOnHF = value;
            return this;
        }
    }
}
