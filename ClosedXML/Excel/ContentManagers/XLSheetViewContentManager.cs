using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ClosedXML.Excel.ContentManagers
{
    internal enum XLSheetViewContents
    {
        Pane,
        Selection,
        PivotSelection,
        ExtensionList
    }

    internal class XLSheetViewContentManager : XLBaseContentManager<XLSheetViewContents>
    {
        public XLSheetViewContentManager(SheetView sheetView)
        {
            contents.Add(XLSheetViewContents.Pane, sheetView.Elements<Pane>().LastOrDefault());
            contents.Add(XLSheetViewContents.Selection, sheetView.Elements<Selection>().LastOrDefault());
            contents.Add(XLSheetViewContents.PivotSelection, sheetView.Elements<PivotSelection>().LastOrDefault());
            contents.Add(XLSheetViewContents.ExtensionList, sheetView.Elements<ExtensionList>().LastOrDefault());
        }
    }
}
