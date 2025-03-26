using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System;
using System.Linq;
using ClosedXML.Utils;
using ClosedXML.Extensions;
using ClosedXML.IO;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Excel.IO;

#nullable disable

internal static class WorksheetPartReader
{
    public static void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws, out PageSetupProperties pageSetupProperties)
    {
        pageSetupProperties = null;
        if (sheetProperty == null) return;

        if (sheetProperty.TabColor != null)
            ws.TabColor = sheetProperty.TabColor.ToClosedXMLColor();

        if (sheetProperty.OutlineProperties != null)
        {
            if (sheetProperty.OutlineProperties.SummaryBelow != null)
            {
                ws.Outline.SummaryVLocation = sheetProperty.OutlineProperties.SummaryBelow
                    ? XLOutlineSummaryVLocation.Bottom
                    : XLOutlineSummaryVLocation.Top;
            }

            if (sheetProperty.OutlineProperties.SummaryRight != null)
            {
                ws.Outline.SummaryHLocation = sheetProperty.OutlineProperties.SummaryRight
                    ? XLOutlineSummaryHLocation.Right
                    : XLOutlineSummaryHLocation.Left;
            }
        }

        if (sheetProperty.PageSetupProperties != null)
            pageSetupProperties = sheetProperty.PageSetupProperties;
    }

    public static void LoadSheetViews(SheetViews sheetViews, XLWorksheet ws)
    {
        if (sheetViews == null) return;

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();

        if (sheetView == null) return;

        if (sheetView.RightToLeft != null) ws.RightToLeft = sheetView.RightToLeft.Value;
        if (sheetView.ShowFormulas != null) ws.ShowFormulas = sheetView.ShowFormulas.Value;
        if (sheetView.ShowGridLines != null) ws.ShowGridLines = sheetView.ShowGridLines.Value;
        if (sheetView.ShowOutlineSymbols != null)
            ws.ShowOutlineSymbols = sheetView.ShowOutlineSymbols.Value;
        if (sheetView.ShowRowColHeaders != null) ws.ShowRowColHeaders = sheetView.ShowRowColHeaders.Value;
        if (sheetView.ShowRuler != null) ws.ShowRuler = sheetView.ShowRuler.Value;
        if (sheetView.ShowWhiteSpace != null) ws.ShowWhiteSpace = sheetView.ShowWhiteSpace.Value;
        if (sheetView.ShowZeros != null) ws.ShowZeros = sheetView.ShowZeros.Value;
        if (sheetView.TabSelected != null) ws.TabSelected = sheetView.TabSelected.Value;

        var selection = sheetView.Elements<Selection>().FirstOrDefault();
        if (selection != null)
        {
            if (selection.SequenceOfReferences != null)
                ws.Ranges(selection.SequenceOfReferences.InnerText.Replace(" ", ",")).Select();

            if (selection.ActiveCell != null)
                ws.Cell(selection.ActiveCell).SetActive();
        }

        if (sheetView.ZoomScale != null)
            ws.SheetView.ZoomScale = (int)UInt32Value.ToUInt32(sheetView.ZoomScale);
        if (sheetView.ZoomScaleNormal != null)
            ws.SheetView.ZoomScaleNormal = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleNormal);
        if (sheetView.ZoomScalePageLayoutView != null)
            ws.SheetView.ZoomScalePageLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScalePageLayoutView);
        if (sheetView.ZoomScaleSheetLayoutView != null)
            ws.SheetView.ZoomScaleSheetLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleSheetLayoutView);

        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (new[] { PaneStateValues.Frozen, PaneStateValues.FrozenSplit }.Contains(pane?.State?.Value ?? PaneStateValues.Split))
        {
            if (pane.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (Int32)pane.HorizontalSplit.Value;
            if (pane.VerticalSplit != null)
                ws.SheetView.SplitRow = (Int32)pane.VerticalSplit.Value;
        }

        if (XLHelper.IsValidA1Address(sheetView.TopLeftCell))
            ws.SheetView.TopLeftCellAddress = ws.Cell(sheetView.TopLeftCell.Value).Address;
    }

    public static void LoadHyperlinks(Hyperlinks hyperlinks, WorksheetPart worksheetPart, XLWorksheet ws)
    {
        var hyperlinkDictionary = new Dictionary<String, Uri>();
        if (worksheetPart.HyperlinkRelationships != null)
            hyperlinkDictionary = worksheetPart.HyperlinkRelationships.ToDictionary(hr => hr.Id, hr => hr.Uri);

        if (hyperlinks == null) return;

        foreach (Hyperlink hl in hyperlinks.Elements<Hyperlink>())
        {
            if (hl.Reference.Value.Equals("#REF")) continue;
            String tooltip = hl.Tooltip != null ? hl.Tooltip.Value : String.Empty;
            var xlRange = ws.Range(hl.Reference.Value);
            foreach (XLCell xlCell in xlRange.Cells())
            {
                if (hl.Id != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hyperlinkDictionary[hl.Id], tooltip));
                else if (hl.Location != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Location.Value, tooltip));
                else
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Reference.Value, tooltip));
            }
        }
    }
}
