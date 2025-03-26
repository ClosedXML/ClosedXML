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
}
