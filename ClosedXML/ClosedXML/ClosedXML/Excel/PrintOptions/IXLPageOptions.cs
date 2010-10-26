using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    
    public interface IXLPageOptions
    {
        IXLPrintAreas PrintAreas { get; }
        Int32 FirstRowToRepeatAtTop { get; }
        Int32 LastRowToRepeatAtTop { get; }
        void SetRowsToRepeatAtTop(String range);
        void SetRowsToRepeatAtTop(Int32 firstRowToRepeatAtTop, Int32 lastRowToRepeatAtTop);
        Int32 FirstColumnToRepeatAtLeft { get; }
        Int32 LastColumnToRepeatAtLeft { get; }
        void SetColumnsToRepeatAtLeft(Int32 firstColumnToRepeatAtLeft, Int32 lastColumnToRepeatAtLeft);
        void SetColumnsToRepeatAtLeft(String range);
        XLPageOrientation PageOrientation { get; set; }
        Int32 PagesWide { get; set; }
        Int32 PagesTall { get; set; }
        Int32 Scale { get; set; }
        Int32 HorizontalDpi { get; set; }
        Int32 VerticalDpi { get; set; }
        Int32 FirstPageNumber { get; set; }
        Boolean CenterHorizontally { get; set; }
        Boolean CenterVertically { get; set; }
        void AdjustTo(Int32 pctOfNormalSize);
        void FitToPages(Int32 pagesWide, Int32 pagesTall);
        XLPaperSize PaperSize { get; set; }
        IXLMargins Margins { get; }

        IXLHeaderFooter Header { get; }
        IXLHeaderFooter Footer { get; }
        Boolean ScaleHFWithDocument { get; set; }
        Boolean AlignHFWithMargins { get; set; }

        Boolean ShowGridlines { get; set; }
        Boolean ShowRowAndColumnHeadings { get; set; }
        Boolean BlackAndWhite { get; set; }
        Boolean DraftQuality { get; set; }
        XLPageOrderValues PageOrder { get; set; }
        XLShowCommentsValues ShowComments { get; set; }


        List<Int32> RowBreaks { get; }
        List<Int32> ColumnBreaks { get; }
        void AddHorizontalPageBreak(Int32 row);
        void AddVerticalPageBreak(Int32 column);

        XLPrintErrorValues PrintErrorValue { get; set; }

    }
}
