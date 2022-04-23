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
                ColumnBreaks = new List<int>();
                RowBreaks = new List<int>();
            }
        }
        public IXLPrintAreas PrintAreas { get; private set; }


        public int FirstRowToRepeatAtTop { get; private set; }
        public int LastRowToRepeatAtTop { get; private set; }
        public void SetRowsToRepeatAtTop(string range)
        {
            var arrRange = range.Replace("$", "").Split(':');
            SetRowsToRepeatAtTop(int.Parse(arrRange[0]), int.Parse(arrRange[1]));
        }
        public void SetRowsToRepeatAtTop(int firstRowToRepeatAtTop, int lastRowToRepeatAtTop)
        {
            if (firstRowToRepeatAtTop <= 0) throw new ArgumentOutOfRangeException("The first row has to be greater than zero.");
            if (firstRowToRepeatAtTop > lastRowToRepeatAtTop) throw new ArgumentOutOfRangeException("The first row has to be less than the second row.");

            FirstRowToRepeatAtTop = firstRowToRepeatAtTop;
            LastRowToRepeatAtTop = lastRowToRepeatAtTop;
        }
        public int FirstColumnToRepeatAtLeft { get; private set; }
        public int LastColumnToRepeatAtLeft { get; private set; }
        public void SetColumnsToRepeatAtLeft(string range)
        {
            var arrRange = range.Replace("$", "").Split(':');
            if (int.TryParse(arrRange[0], out var iTest))
                SetColumnsToRepeatAtLeft(int.Parse(arrRange[0]), int.Parse(arrRange[1]));
            else
                SetColumnsToRepeatAtLeft(arrRange[0], arrRange[1]);
        }
        public void SetColumnsToRepeatAtLeft(string firstColumnToRepeatAtLeft, string lastColumnToRepeatAtLeft)
        {
            SetColumnsToRepeatAtLeft(XLHelper.GetColumnNumberFromLetter(firstColumnToRepeatAtLeft), XLHelper.GetColumnNumberFromLetter(lastColumnToRepeatAtLeft));
        }
        public void SetColumnsToRepeatAtLeft(int firstColumnToRepeatAtLeft, int lastColumnToRepeatAtLeft)
        {
            if (firstColumnToRepeatAtLeft <= 0) throw new ArgumentOutOfRangeException("The first column has to be greater than zero.");
            if (firstColumnToRepeatAtLeft > lastColumnToRepeatAtLeft) throw new ArgumentOutOfRangeException("The first column has to be less than the second column.");

            FirstColumnToRepeatAtLeft = firstColumnToRepeatAtLeft;
            LastColumnToRepeatAtLeft = lastColumnToRepeatAtLeft;
        }

        public XLPageOrientation PageOrientation { get; set; }
        public XLPaperSize PaperSize { get; set; }
        public int HorizontalDpi { get; set; }
        public int VerticalDpi { get; set; }
        public uint? FirstPageNumber { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public XLPrintErrorValues PrintErrorValue { get; set; }
        public IXLMargins Margins { get; set; }

        private int _pagesWide;
        public int PagesWide
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

        private int _pagesTall;
        public int PagesTall
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

        private int _scale;
        public int Scale
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

        public void AdjustTo(int percentageOfNormalSize)
        {
            Scale = percentageOfNormalSize;
            _pagesWide = 0;
            _pagesTall = 0;
        }
        public void FitToPages(int pagesWide, int pagesTall)
        {
            _pagesWide = pagesWide;
            _pagesTall = pagesTall;
            _scale = 0;
        }


        public IXLHeaderFooter Header { get; private set; }
        public IXLHeaderFooter Footer { get; private set; }

        public bool ScaleHFWithDocument { get; set; }
        public bool AlignHFWithMargins { get; set; }

        public bool ShowGridlines { get; set; }
        public bool ShowRowAndColumnHeadings { get; set; }
        public bool BlackAndWhite { get; set; }
        public bool DraftQuality { get; set; }

        public XLPageOrderValues PageOrder { get; set; }
        public XLShowCommentsValues ShowComments { get; set; }

        public List<int> RowBreaks { get; private set; }
        public List<int> ColumnBreaks { get; private set; }
        public void AddHorizontalPageBreak(int row)
        {
            if (!RowBreaks.Contains(row))
                RowBreaks.Add(row);
            RowBreaks.Sort();
        }
        public void AddVerticalPageBreak(int column)
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
        public IXLPageSetup SetPagesWide(int value) { PagesWide = value; return this; }
        public IXLPageSetup SetPagesTall(int value) { PagesTall = value; return this; }
        public IXLPageSetup SetScale(int value) { Scale = value; return this; }
        public IXLPageSetup SetHorizontalDpi(int value) { HorizontalDpi = value; return this; }
        public IXLPageSetup SetVerticalDpi(int value) { VerticalDpi = value; return this; }
        public IXLPageSetup SetFirstPageNumber(uint? value) { FirstPageNumber = value; return this; }
        public IXLPageSetup SetCenterHorizontally() { CenterHorizontally = true; return this; }	public IXLPageSetup SetCenterHorizontally(bool value) { CenterHorizontally = value; return this; }
        public IXLPageSetup SetCenterVertically() { CenterVertically = true; return this; }	public IXLPageSetup SetCenterVertically(bool value) { CenterVertically = value; return this; }
        public IXLPageSetup SetPaperSize(XLPaperSize value) { PaperSize = value; return this; }
        public IXLPageSetup SetScaleHFWithDocument() { ScaleHFWithDocument = true; return this; }	public IXLPageSetup SetScaleHFWithDocument(bool value) { ScaleHFWithDocument = value; return this; }
        public IXLPageSetup SetAlignHFWithMargins() { AlignHFWithMargins = true; return this; }	public IXLPageSetup SetAlignHFWithMargins(bool value) { AlignHFWithMargins = value; return this; }
        public IXLPageSetup SetShowGridlines() { ShowGridlines = true; return this; }	public IXLPageSetup SetShowGridlines(bool value) { ShowGridlines = value; return this; }
        public IXLPageSetup SetShowRowAndColumnHeadings() { ShowRowAndColumnHeadings = true; return this; }	public IXLPageSetup SetShowRowAndColumnHeadings(bool value) { ShowRowAndColumnHeadings = value; return this; }
        public IXLPageSetup SetBlackAndWhite() { BlackAndWhite = true; return this; }	public IXLPageSetup SetBlackAndWhite(bool value) { BlackAndWhite = value; return this; }
        public IXLPageSetup SetDraftQuality() { DraftQuality = true; return this; }	public IXLPageSetup SetDraftQuality(bool value) { DraftQuality = value; return this; }
        public IXLPageSetup SetPageOrder(XLPageOrderValues value) { PageOrder = value; return this; }
        public IXLPageSetup SetShowComments(XLShowCommentsValues value) { ShowComments = value; return this; }
        public IXLPageSetup SetPrintErrorValue(XLPrintErrorValues value) { PrintErrorValue = value; return this; }

        public bool DifferentFirstPageOnHF { get; set; }
        public IXLPageSetup SetDifferentFirstPageOnHF()
        {
            return SetDifferentFirstPageOnHF(true);
        }
        public IXLPageSetup SetDifferentFirstPageOnHF(bool value)
        {
            DifferentFirstPageOnHF = value;
            return this;
        }
        public bool DifferentOddEvenPagesOnHF { get; set; }
        public IXLPageSetup SetDifferentOddEvenPagesOnHF()
        {
            return SetDifferentOddEvenPagesOnHF(true);
        }
        public IXLPageSetup SetDifferentOddEvenPagesOnHF(bool value)
        {
            DifferentOddEvenPagesOnHF = value;
            return this;
        }
    }
}
