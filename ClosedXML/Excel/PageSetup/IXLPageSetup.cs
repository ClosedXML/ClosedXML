﻿using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLPageOrientation { Default, Portrait, Landscape }
    public enum XLPaperSize
    {
        LetterPaper = 1,
        LetterSmallPaper = 2,
        TabloidPaper = 3,
        LedgerPaper = 4,
        LegalPaper = 5,
        StatementPaper = 6,
        ExecutivePaper = 7,
        A3Paper = 8,
        A4Paper = 9,
        A4SmallPaper = 10,
        A5Paper = 11,
        B4Paper = 12,
        B5Paper = 13,
        FolioPaper = 14,
        QuartoPaper = 15,
        StandardPaper = 16,
        StandardPaper1 = 17,
        NotePaper = 18,
        No9Envelope = 19,
        No10Envelope = 20,
        No11Envelope = 21,
        No12Envelope = 22,
        No14Envelope = 23,
        CPaper = 24,
        DPaper = 25,
        EPaper = 26,
        DlEnvelope = 27,
        C5Envelope = 28,
        C3Envelope = 29,
        C4Envelope = 30,
        C6Envelope = 31,
        C65Envelope = 32,
        B4Envelope = 33,
        B5Envelope = 34,
        B6Envelope = 35,
        ItalyEnvelope = 36,
        MonarchEnvelope = 37,
        No634Envelope = 38,
        UsStandardFanfold = 39,
        GermanStandardFanfold = 40,
        GermanLegalFanfold = 41,
        IsoB4 = 42,
        JapaneseDoublePostcard = 43,
        StandardPaper2 = 44,
        StandardPaper3 = 45,
        StandardPaper4 = 46,
        InviteEnvelope = 47,
        LetterExtraPaper = 50,
        LegalExtraPaper = 51,
        TabloidExtraPaper = 52,
        A4ExtraPaper = 53,
        LetterTransversePaper = 54,
        A4TransversePaper = 55,
        LetterExtraTransversePaper = 56,
        SuperaSuperaA4Paper = 57,
        SuperbSuperbA3Paper = 58,
        LetterPlusPaper = 59,
        A4PlusPaper = 60,
        A5TransversePaper = 61,
        JisB5TransversePaper = 62,
        A3ExtraPaper = 63,
        A5ExtraPaper = 64,
        IsoB5ExtraPaper = 65,
        A2Paper = 66,
        A3TransversePaper = 67,
        A3ExtraTransversePaper = 68
    }
    public enum XLPageOrderValues { DownThenOver, OverThenDown }
    public enum XLShowCommentsValues { None, AtEnd, AsDisplayed }
    public enum XLPrintErrorValues { Blank, Dash, Displayed, NA }
    
    public interface IXLPageSetup
    {
        /// <summary>
        /// Gets an object to manage the print areas of the worksheet.
        /// </summary>
        IXLPrintAreas PrintAreas { get; }
        /// <summary>
        /// Gets the first row that will repeat on the top of the printed pages.
        /// <para>Use SetRowsToRepeatAtTop() to set the rows that will be repeated on the top of the printed pages.</para>
        /// </summary>
        Int32 FirstRowToRepeatAtTop { get; }
        /// <summary>
        /// Gets the last row that will repeat on the top of the printed pages.
        /// <para>Use SetRowsToRepeatAtTop() to set the rows that will be repeated on the top of the printed pages.</para>
        /// </summary>
        Int32 LastRowToRepeatAtTop { get; }
        /// <summary>
        /// Sets the rows to repeat on the top of the printed pages.
        /// </summary>
        /// <param name="range">The range of rows to repeat on the top of the printed pages.</param>
        void SetRowsToRepeatAtTop(String range);
        /// <summary>
        /// Sets the rows to repeat on the top of the printed pages.
        /// </summary>
        /// <param name="firstRowToRepeatAtTop">The first row to repeat at top.</param>
        /// <param name="lastRowToRepeatAtTop">The last row to repeat at top.</param>
        void SetRowsToRepeatAtTop(Int32 firstRowToRepeatAtTop, Int32 lastRowToRepeatAtTop);

        /// <summary>Gets the first column to repeat on the left of the printed pages.</summary>
        /// <value>The first column to repeat on the left of the printed pages.</value>
        Int32 FirstColumnToRepeatAtLeft { get; }
        /// <summary>Gets the last column to repeat on the left of the printed pages.</summary>
        /// <value>The last column to repeat on the left of the printed pages.</value>
        Int32 LastColumnToRepeatAtLeft { get; }
        /// <summary>
        /// Sets the rows to repeat on the left of the printed pages.
        /// </summary>
        /// <param name="firstColumnToRepeatAtLeft">The first column to repeat at left.</param>
        /// <param name="lastColumnToRepeatAtLeft">The last column to repeat at left.</param>
        void SetColumnsToRepeatAtLeft(Int32 firstColumnToRepeatAtLeft, Int32 lastColumnToRepeatAtLeft);
        /// <summary>
        /// Sets the rows to repeat on the left of the printed pages.
        /// </summary>
        /// <param name="range">The range of rows to repeat on the left of the printed pages.</param>
        void SetColumnsToRepeatAtLeft(String range);

        /// <summary>Gets or sets the page orientation for printing.</summary>
        /// <value>The page orientation.</value>
        XLPageOrientation PageOrientation { get; set; }
        /// <summary>
        /// Gets or sets the number of pages wide (horizontal) the worksheet will be printed on.
        /// <para>If you don't specify the PagesTall, Excel will adjust that value</para>
        /// <para>based on the contents of the worksheet and the PagesWide number.</para>
        /// <para>Setting this value will override the Scale value.</para>
        /// </summary>
        Int32 PagesWide { get; set; }
        /// <summary>
        /// Gets or sets the number of pages tall (vertical) the worksheet will be printed on.
        /// <para>If you don't specify the PagesWide, Excel will adjust that value</para>
        /// <para>based on the contents of the worksheet and the PagesTall number.</para>
        /// <para>Setting this value will override the Scale value.</para>
        /// </summary>
        Int32 PagesTall { get; set; }
        /// <summary>
        /// Gets or sets the scale at which the worksheet will be printed.
        /// <para>The worksheet will be printed on as many pages as necessary to print at the given scale.</para>
        /// <para>Setting this value will override the PagesWide and PagesTall values.</para>
        /// </summary>
        Int32 Scale { get; set; }
        /// <summary>
        /// Gets or sets the horizontal dpi for printing the worksheet.
        /// </summary>
        Int32 HorizontalDpi { get; set; }
        /// <summary>
        /// Gets or sets the vertical dpi for printing the worksheet.
        /// </summary>
        Int32 VerticalDpi { get; set; }
        /// <summary>
        /// Gets or sets the page number that will begin the printout.
        /// <para>For example, the first page of your printout could be numbered page 5.</para>
        /// </summary>
        Int64 FirstPageNumber { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the worksheet will be centered on the page horizontally.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the worksheet will be centered on the page horizontally; otherwise, <c>false</c>.
        /// </value>
        Boolean CenterHorizontally { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the worksheet will be centered on the page vertically.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the worksheet will be centered on the page vartically; otherwise, <c>false</c>.
        /// </value>
        Boolean CenterVertically { get; set; }
        /// <summary>
        /// Sets the scale at which the worksheet will be printed. This is equivalent to setting the Scale property.
        /// <para>The worksheet will be printed on as many pages as necessary to print at the given scale.</para>
        /// <para>Setting this value will override the PagesWide and PagesTall values.</para>
        /// </summary>
        /// <param name="pctOfNormalSize">The scale at which the worksheet will be printed.</param>
        void AdjustTo(Int32 pctOfNormalSize);
        /// <summary>
        /// Gets or sets the number of pages the worksheet will be printed on.
        /// <para>This is equivalent to setting both PagesWide and PagesTall properties.</para>
        /// <para>Setting this value will override the Scale value.</para>
        /// </summary>
        /// <param name="pagesWide">The pages wide.</param>
        /// <param name="pagesTall">The pages tall.</param>
        void FitToPages(Int32 pagesWide, Int32 pagesTall);
        /// <summary>
        /// Gets or sets the size of the paper to print the worksheet.
        /// </summary>
        XLPaperSize PaperSize { get; set; }
        /// <summary>
        /// Gets an object to work with the page margins.
        /// </summary>
        IXLMargins Margins { get; }

        /// <summary>
        /// Gets an object to work with the page headers.
        /// </summary>
        IXLHeaderFooter Header { get; }
        /// <summary>
        /// Gets an object to work with the page footers.
        /// </summary>
        IXLHeaderFooter Footer { get; }
        /// <summary>
        /// Gets or sets a value indicating whether Excel will automatically adjust the font size to the scale of the worksheet.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if Excel will automatically adjust the font size to the scale of the worksheet; otherwise, <c>false</c>.
        /// </value>
        Boolean ScaleHFWithDocument { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the header and footer margins are aligned with the left and right margins of the worksheet.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the header and footer margins are aligned with the left and right margins of the worksheet; otherwise, <c>false</c>.
        /// </value>
        Boolean AlignHFWithMargins { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the gridlines will be printed.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the gridlines will be printed; otherwise, <c>false</c>.
        /// </value>
        Boolean ShowGridlines { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether to show row numbers and column letters/numbers.
        /// </summary>
        /// <value>
        /// 	<c>true</c> to show row numbers and column letters/numbers; otherwise, <c>false</c>.
        /// </value>
        Boolean ShowRowAndColumnHeadings { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the worksheet will be printed in black and white.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the worksheet will be printed in black and white; otherwise, <c>false</c>.
        /// </value>
        Boolean BlackAndWhite { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the worksheet will be printed in draft quality.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the worksheet will be printed in draft quality; otherwise, <c>false</c>.
        /// </value>
        Boolean DraftQuality { get; set; }
        /// <summary>
        /// Gets or sets the page order for printing.
        /// </summary>
        XLPageOrderValues PageOrder { get; set; }
        /// <summary>
        /// Gets or sets how the comments will be printed.
        /// </summary>
        XLShowCommentsValues ShowComments { get; set; }

        /// <summary>
        /// Gets a list with the row breaks (for printing).
        /// </summary>
        List<Int32> RowBreaks { get; }
        /// <summary>
        /// Gets a list with the column breaks (for printing).
        /// </summary>
        List<Int32> ColumnBreaks { get; }
        /// <summary>
        /// Adds a horizontal page break after the given row.
        /// </summary>
        /// <param name="row">The row to insert the break.</param>
        void AddHorizontalPageBreak(Int32 row);

        /// <summary>
        /// Adds a vertical page break after the given column.
        /// </summary>
        /// <param name="column">The column to insert the break.</param>
        void AddVerticalPageBreak(Int32 column);

        /// <summary>
        /// Gets or sets how error values will be printed.
        /// </summary>
        XLPrintErrorValues PrintErrorValue { get; set; }

        IXLPageSetup SetPageOrientation(XLPageOrientation value);
        IXLPageSetup SetPagesWide(Int32 value);
        IXLPageSetup SetPagesTall(Int32 value);
        IXLPageSetup SetScale(Int32 value);
        IXLPageSetup SetHorizontalDpi(Int32 value);
        IXLPageSetup SetVerticalDpi(Int32 value);
        IXLPageSetup SetFirstPageNumber(Int64 value);
        IXLPageSetup SetCenterHorizontally(); IXLPageSetup SetCenterHorizontally(Boolean value);
        IXLPageSetup SetCenterVertically(); IXLPageSetup SetCenterVertically(Boolean value);
        IXLPageSetup SetPaperSize(XLPaperSize value);
        IXLPageSetup SetScaleHFWithDocument(); IXLPageSetup SetScaleHFWithDocument(Boolean value);
        IXLPageSetup SetAlignHFWithMargins(); IXLPageSetup SetAlignHFWithMargins(Boolean value);
        IXLPageSetup SetShowGridlines(); IXLPageSetup SetShowGridlines(Boolean value);
        IXLPageSetup SetShowRowAndColumnHeadings(); IXLPageSetup SetShowRowAndColumnHeadings(Boolean value);
        IXLPageSetup SetBlackAndWhite(); IXLPageSetup SetBlackAndWhite(Boolean value);
        IXLPageSetup SetDraftQuality(); IXLPageSetup SetDraftQuality(Boolean value);
        IXLPageSetup SetPageOrder(XLPageOrderValues value);
        IXLPageSetup SetShowComments(XLShowCommentsValues value);
        IXLPageSetup SetPrintErrorValue(XLPrintErrorValues value);


        Boolean DifferentFirstPageOnHF { get; set; }
        IXLPageSetup SetDifferentFirstPageOnHF(); IXLPageSetup SetDifferentFirstPageOnHF(Boolean value);
        Boolean DifferentOddEvenPagesOnHF { get; set; }
        IXLPageSetup SetDifferentOddEvenPagesOnHF(); IXLPageSetup SetDifferentOddEvenPagesOnHF(Boolean value);
    }
}
