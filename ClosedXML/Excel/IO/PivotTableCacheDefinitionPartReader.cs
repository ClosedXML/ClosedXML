#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Cells;
using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.IO
{
    internal class PivotTableCacheDefinitionPartReader
    {
        internal static void Load(WorkbookPart workbookPart, Sheets sheets, XLWorkbook workbook, Dictionary<int, DifferentialFormat> differentialFormats)
        {
            foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                if (pivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.WorksheetSource != null)
                {
                    var pivotSourceReference = ParsePivotSourceReference(pivotTableCacheDefinitionPart, workbook);
                    if (pivotSourceReference == null)
                        // We don't support external sources
                        continue;

                    var pivotCache = workbook.PivotCachesInternal.Add(pivotSourceReference);

                    // If WorkbookCacheRelId already has a value, it means the pivot source is being reused
                    if (string.IsNullOrWhiteSpace(pivotCache.WorkbookCacheRelId))
                    {
                        pivotCache.WorkbookCacheRelId = workbookPart.GetIdOfPart(pivotTableCacheDefinitionPart);
                    }

                    var cacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;
                    if (cacheDefinition.MissingItemsLimit is not null)
                    {
                        if (cacheDefinition.MissingItemsLimit == 0U)
                        {
                            pivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
                        }
                        else if (cacheDefinition.MissingItemsLimit == XLHelper.MaxRowNumber)
                        {
                            pivotCache.ItemsToRetainPerField = XLItemsToRetain.Max;
                        }
                    }

                    if (pivotTableCacheDefinitionPart.PivotCacheDefinition?.CacheFields is { } cacheFields)
                    {
                        ReadCacheFields(cacheFields, pivotCache);
                        if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } recordsPart)
                        {
                            ReadRecords(recordsPart, pivotCache);
                        }
                    }

                    if (pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData != null)
                    {
                        pivotCache.SaveSourceData = pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData.Value;
                    }
                }
            }

            // Delay loading of pivot tables until all sheets have been loaded
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (worksheetPart is not null)
                {
                    var ws = (XLWorksheet)workbook.WorksheetsInternal.Worksheet(dSheet.Name);

                    foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                    {
                        var cache = pivotTablePart.PivotTableCacheDefinitionPart;
                        var cacheDefinitionRelId = workbookPart.GetIdOfPart(cache);

                        var pivotSource = workbook.PivotCachesInternal
                            .FirstOrDefault<XLPivotCache>(ps => ps.WorkbookCacheRelId == cacheDefinitionRelId);

                        if (pivotSource == null)
                        {
                            // If it's missing, find a 'similar' pivot cache, i.e. one that's based on the same source range/table
                            pivotSource = workbook.PivotCachesInternal
                                .FirstOrDefault<XLPivotCache>(ps =>
                                    ps.PivotSourceReference.Equals(ParsePivotSourceReference(cache, workbook)));
                        }

                        var pivotTableDefinition = pivotTablePart.PivotTableDefinition;

                        var target = ws.FirstCell();
                        if (pivotTableDefinition?.Location?.Reference?.HasValue ?? false)
                        {
                            ws.Range(pivotTableDefinition.Location.Reference.Value).Clear(XLClearOptions.All);
                            target = ws.Range(pivotTableDefinition.Location.Reference.Value).FirstCell();
                        }

                        if (target != null && pivotSource != null)
                        {
                            var pt = ws.PivotTables.Add(pivotTableDefinition.Name, target, pivotSource) as XLPivotTable;

                            if (!String.IsNullOrWhiteSpace(
                                    StringValue.ToString(pivotTableDefinition?.ColumnHeaderCaption ?? String.Empty)))
                                pt.SetColumnHeaderCaption(StringValue.ToString(pivotTableDefinition.ColumnHeaderCaption));

                            if (!String.IsNullOrWhiteSpace(
                                    StringValue.ToString(pivotTableDefinition?.RowHeaderCaption ?? String.Empty)))
                                pt.SetRowHeaderCaption(StringValue.ToString(pivotTableDefinition.RowHeaderCaption));

                            pt.RelId = worksheetPart.GetIdOfPart(pivotTablePart);
                            pt.CacheDefinitionRelId = pivotTablePart.GetIdOfPart(cache);

                            if (pivotTableDefinition.MergeItem != null)
                                pt.MergeAndCenterWithLabels = pivotTableDefinition.MergeItem.Value;
                            if (pivotTableDefinition.Indent != null) pt.RowLabelIndent = (int)pivotTableDefinition.Indent.Value;
                            if (pivotTableDefinition.PageOverThenDown != null)
                                pt.FilterAreaOrder = pivotTableDefinition.PageOverThenDown.Value
                                    ? XLFilterAreaOrder.OverThenDown
                                    : XLFilterAreaOrder.DownThenOver;
                            if (pivotTableDefinition.PageWrap != null)
                                pt.FilterFieldsPageWrap = (int)pivotTableDefinition.PageWrap.Value;
                            if (pivotTableDefinition.UseAutoFormatting != null)
                                pt.AutofitColumns = pivotTableDefinition.UseAutoFormatting.Value;
                            if (pivotTableDefinition.PreserveFormatting != null)
                                pt.PreserveCellFormatting = pivotTableDefinition.PreserveFormatting.Value;
                            if (pivotTableDefinition.RowGrandTotals != null)
                                pt.ShowGrandTotalsRows = pivotTableDefinition.RowGrandTotals.Value;
                            if (pivotTableDefinition.ColumnGrandTotals != null)
                                pt.ShowGrandTotalsColumns = pivotTableDefinition.ColumnGrandTotals.Value;
                            if (pivotTableDefinition.SubtotalHiddenItems != null)
                                pt.FilteredItemsInSubtotals = pivotTableDefinition.SubtotalHiddenItems.Value;
                            if (pivotTableDefinition.MultipleFieldFilters != null)
                                pt.AllowMultipleFilters = pivotTableDefinition.MultipleFieldFilters.Value;
                            if (pivotTableDefinition.CustomListSort != null)
                                pt.UseCustomListsForSorting = pivotTableDefinition.CustomListSort.Value;
                            if (pivotTableDefinition.ShowDrill != null)
                                pt.ShowExpandCollapseButtons = pivotTableDefinition.ShowDrill.Value;
                            if (pivotTableDefinition.ShowDataTips != null)
                                pt.ShowContextualTooltips = pivotTableDefinition.ShowDataTips.Value;
                            if (pivotTableDefinition.ShowMemberPropertyTips != null)
                                pt.ShowPropertiesInTooltips = pivotTableDefinition.ShowMemberPropertyTips.Value;
                            if (pivotTableDefinition.ShowHeaders != null)
                                pt.DisplayCaptionsAndDropdowns = pivotTableDefinition.ShowHeaders.Value;
                            if (pivotTableDefinition.GridDropZones != null)
                                pt.ClassicPivotTableLayout = pivotTableDefinition.GridDropZones.Value;
                            if (pivotTableDefinition.ShowEmptyRow != null)
                                pt.ShowEmptyItemsOnRows = pivotTableDefinition.ShowEmptyRow.Value;
                            if (pivotTableDefinition.ShowEmptyColumn != null)
                                pt.ShowEmptyItemsOnColumns = pivotTableDefinition.ShowEmptyColumn.Value;
                            if (pivotTableDefinition.ShowItems != null)
                                pt.DisplayItemLabels = pivotTableDefinition.ShowItems.Value;
                            if (pivotTableDefinition.FieldListSortAscending != null)
                                pt.SortFieldsAtoZ = pivotTableDefinition.FieldListSortAscending.Value;
                            if (pivotTableDefinition.PrintDrill != null)
                                pt.PrintExpandCollapsedButtons = pivotTableDefinition.PrintDrill.Value;
                            if (pivotTableDefinition.ItemPrintTitles != null)
                                pt.RepeatRowLabels = pivotTableDefinition.ItemPrintTitles.Value;
                            if (pivotTableDefinition.FieldPrintTitles != null)
                                pt.PrintTitles = pivotTableDefinition.FieldPrintTitles.Value;
                            if (pivotTableDefinition.EnableDrill != null)
                                pt.EnableShowDetails = pivotTableDefinition.EnableDrill.Value;

                            if (pivotTableDefinition.ShowMissing != null && pivotTableDefinition.MissingCaption != null)
                                pt.EmptyCellReplacement = pivotTableDefinition.MissingCaption.Value;

                            if (pivotTableDefinition.ShowError != null && pivotTableDefinition.ErrorCaption != null)
                                pt.ErrorValueReplacement = pivotTableDefinition.ErrorCaption.Value;

                            var pivotTableDefinitionExtensionList =
                                pivotTableDefinition.GetFirstChild<PivotTableDefinitionExtensionList>();
                            var pivotTableDefinitionExtension =
                                pivotTableDefinitionExtensionList?.GetFirstChild<PivotTableDefinitionExtension>();
                            var pivotTableDefinition2 = pivotTableDefinitionExtension
                                ?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition>();
                            if (pivotTableDefinition2 != null)
                            {
                                if (pivotTableDefinition2.EnableEdit != null)
                                    pt.EnableCellEditing = pivotTableDefinition2.EnableEdit.Value;
                                if (pivotTableDefinition2.HideValuesRow != null)
                                    pt.ShowValuesRow = !pivotTableDefinition2.HideValuesRow.Value;
                            }

                            var pivotTableStyle = pivotTableDefinition.GetFirstChild<PivotTableStyle>();
                            if (pivotTableStyle != null)
                            {
                                if (pivotTableStyle.Name != null)
                                    pt.Theme = (XLPivotTableTheme)Enum.Parse(typeof(XLPivotTableTheme), pivotTableStyle.Name);
                                else
                                    pt.Theme = XLPivotTableTheme.None;

                                pt.ShowRowHeaders = OpenXmlHelper.GetBooleanValueAsBool(pivotTableStyle.ShowRowHeaders, false);
                                pt.ShowColumnHeaders =
                                    OpenXmlHelper.GetBooleanValueAsBool(pivotTableStyle.ShowColumnHeaders, false);
                                pt.ShowRowStripes = OpenXmlHelper.GetBooleanValueAsBool(pivotTableStyle.ShowRowStripes, false);
                                pt.ShowColumnStripes =
                                    OpenXmlHelper.GetBooleanValueAsBool(pivotTableStyle.ShowColumnStripes, false);
                            }

                            // Subtotal configuration
                            if (pivotTableDefinition.PivotFields.Cast<PivotField>().All(pf =>
                                    (pf.DefaultSubtotal == null || pf.DefaultSubtotal.Value)
                                    && (pf.SubtotalTop == null || pf.SubtotalTop == true)))
                                pt.SetSubtotals(XLPivotSubtotals.AtTop);
                            else if (pivotTableDefinition.PivotFields.Cast<PivotField>().All(pf =>
                                         (pf.DefaultSubtotal == null || pf.DefaultSubtotal.Value)
                                         && (pf.SubtotalTop != null && pf.SubtotalTop.Value == false)))
                                pt.SetSubtotals(XLPivotSubtotals.AtBottom);
                            else
                                pt.SetSubtotals(XLPivotSubtotals.DoNotShow);

                            // Row labels
                            if (pivotTableDefinition.RowFields != null)
                            {
                                foreach (var rf in pivotTableDefinition.RowFields.Cast<Field>())
                                {
                                    if (rf.Index < pivotTableDefinition.PivotFields.Count)
                                    {
                                        if (rf.Index.Value == -2)
                                        {
                                            pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
                                        }
                                        else
                                        {
                                            if (!(pivotTableDefinition.PivotFields.ElementAt(rf.Index.Value) is PivotField pf))
                                                continue;

                                            var cacheFieldName = pivotSource.FieldNames[rf.Index.Value];

                                            var pivotField = pf.Name != null
                                                ? pt.RowLabels.Add(cacheFieldName, pf.Name.Value)
                                                : pt.RowLabels.Add(cacheFieldName);

                                            if (pivotField != null)
                                            {
                                                LoadFieldOptions(pf, pivotField);
                                                LoadSubtotals(pf, pivotField);

                                                if (pf.SortType != null)
                                                {
                                                    pivotField.SetSort((XLPivotSortType)pf.SortType.Value);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // Column labels
                            if (pivotTableDefinition.ColumnFields != null)
                            {
                                foreach (var cf in pivotTableDefinition.ColumnFields.Cast<Field>())
                                {
                                    IXLPivotField pivotField = null;
                                    if (cf.Index.Value == -2)
                                        pivotField = pt.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
                                    else if (cf.Index < pivotTableDefinition.PivotFields.Count)
                                    {
                                        if (!(pivotTableDefinition.PivotFields.ElementAt(cf.Index.Value) is PivotField pf))
                                            continue;

                                        var cacheFieldName = pivotSource.FieldNames[cf.Index.Value];

                                        pivotField = pf.Name != null
                                            ? pt.ColumnLabels.Add(cacheFieldName, pf.Name.Value)
                                            : pt.ColumnLabels.Add(cacheFieldName);

                                        if (pivotField != null)
                                        {
                                            LoadFieldOptions(pf, pivotField);
                                            LoadSubtotals(pf, pivotField);

                                            if (pf.SortType != null)
                                            {
                                                pivotField.SetSort((XLPivotSortType)pf.SortType.Value);
                                            }
                                        }
                                    }
                                }
                            }

                            // Values
                            if (pivotTableDefinition.DataFields != null)
                            {
                                foreach (var df in pivotTableDefinition.DataFields.Cast<DataField>())
                                {
                                    IXLPivotValue pivotValue = null;
                                    if ((int)df.Field.Value == -2)
                                        pivotValue = pt.Values.Add(XLConstants.PivotTable.ValuesSentinalLabel);
                                    else if (df.Field.Value < pivotTableDefinition.PivotFields.Count)
                                    {
                                        if (!(pivotTableDefinition.PivotFields.ElementAt((int)df.Field.Value) is PivotField pf))
                                            continue;

                                        var cacheFieldName = pivotSource.FieldNames[(int)df.Field.Value];

                                        if (pf.Name != null)
                                            pivotValue = pt.Values.Add(pf.Name.Value, df.Name.Value);
                                        else
                                            pivotValue = pt.Values.Add(cacheFieldName, df.Name.Value);

                                        if (df.NumberFormatId != null)
                                            pivotValue.NumberFormat.SetNumberFormatId((int)df.NumberFormatId.Value);
                                        if (df.Subtotal != null)
                                            pivotValue = pivotValue.SetSummaryFormula(df.Subtotal.Value.ToClosedXml());
                                        if (df.ShowDataAs != null)
                                        {
                                            var calculation = df.ShowDataAs.Value.ToClosedXml();
                                            pivotValue = pivotValue.SetCalculation(calculation);
                                        }

                                        if (df.BaseField?.Value != null)
                                        {
                                            pivotValue.BaseFieldName = pt.PivotCache.FieldNames[df.BaseField.Value];

                                            if (df.BaseItem?.Value != null)
                                            {
                                                var items = pt.PivotCache
                                                    .GetFieldValues(df.BaseField.Value)
                                                    .GetCellValues()
                                                    .Distinct(XLCellValueComparer.OrdinalIgnoreCase)
                                                    .ToList();
                                                var bi = (int)df.BaseItem.Value;
                                                if (bi.Between(0, items.Count - 1))
                                                    pivotValue.BaseItemValue = items[(int)df.BaseItem.Value];
                                            }
                                        }
                                    }
                                }
                            }

                            // Filters
                            if (pivotTableDefinition.PageFields != null)
                            {
                                foreach (var pageField in pivotTableDefinition.PageFields.Cast<PageField>())
                                {
                                    if (!(pivotTableDefinition.PivotFields.ElementAt(pageField.Field.Value) is PivotField pf))
                                        continue;

                                    var cacheFieldValues = pivotSource.GetFieldSharedItems(pageField.Field.Value);

                                    var filterName = pf.Name?.Value ?? pivotSource.FieldNames[pageField.Field.Value];

                                    IXLPivotField rf;
                                    if (pageField.Name?.Value != null)
                                        rf = pt.ReportFilters.Add(filterName, pageField.Name.Value);
                                    else
                                        rf = pt.ReportFilters.Add(filterName);

                                    var openXmlItems = new List<Item>();

                                    if (OpenXmlHelper.GetBooleanValueAsBool(pf.MultipleItemSelectionAllowed, false))
                                    {
                                        openXmlItems.AddRange(pf.Items.Cast<Item>());
                                    }
                                    else if ((pageField.Item?.HasValue ?? false)
                                             && pf.Items.Any()
                                             && cacheFieldValues.Count > 0)
                                    {
                                        if (!(pf.Items.ElementAt(Convert.ToInt32(pageField.Item.Value)) is Item item))
                                            continue;

                                        openXmlItems.Add(item);
                                    }

                                    foreach (var item in openXmlItems)
                                    {
                                        if (!OpenXmlHelper.GetBooleanValueAsBool(item.Hidden, false)
                                            && (item.Index?.HasValue ?? false))
                                        {
                                            var value = cacheFieldValues[item.Index.Value];
                                            rf.AddSelectedValue(value);
                                        }
                                    }
                                }

                                pt.TargetCell = pt.TargetCell.CellAbove(pt.ReportFilters.Count() + 1);
                            }

                            LoadPivotStyleFormats(pt, pivotTableDefinition, differentialFormats);
                        }
                    }
                }
            }
        }

        private static XLPivotSourceReference ParsePivotSourceReference(PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart, XLWorkbook workbook)
        {
            // TODO: Implement other sources besides worksheetSource
            // But for now assume names and references point directly to a range
            var wss = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheSource.WorksheetSource;

            if (!String.IsNullOrEmpty(wss.Id))
            {
                var externalRelationship = pivotTableCacheDefinitionPart.ExternalRelationships.FirstOrDefault(er => er.Id.Equals(wss.Id));
                if (externalRelationship?.IsExternal ?? false)
                {
                    // We don't support external sources
                    return null;
                }
            }

            // Source data of pivot cache are from a table or a named range.
            if (wss.Name is not null)
            {
                return new XLPivotSourceReference(wss.Name);
            }

            // Source data of pivot cache are from an area of a workbook.
            if (wss.Reference is not null && wss.Sheet is not null)
            {
                var bookArea = new XLBookArea(wss.Sheet, XLSheetRange.Parse(wss.Reference));
                return new XLPivotSourceReference(bookArea);
            }

            throw PartStructureException.MissingAttribute();
        }

        private static void LoadPivotStyleFormats(XLPivotTable pt, PivotTableDefinition ptd, Dictionary<Int32, DifferentialFormat> differentialFormats)
        {
            if (ptd.Formats == null)
                return;

            foreach (var format in ptd.Formats.OfType<Format>())
            {
                var pivotArea = format.PivotArea;
                if (pivotArea == null)
                    continue;

                var type = pivotArea.Type ?? PivotAreaValues.Normal;
                var dataOnly = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.DataOnly, true);
                var labelOnly = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.LabelOnly, false);

                if (dataOnly && labelOnly)
                    throw new InvalidOperationException("Cannot have dataOnly and labelOnly both set to true.");

                XLPivotStyleFormat styleFormat;

                if (pivotArea.Field == null && !(pivotArea.PivotAreaReferences?.OfType<PivotAreaReference>()?.Any() ?? false))
                {
                    // If the pivot field is null and doesn't have children (references), we assume this format is a grand total
                    // Example:
                    // <x:pivotArea type="normal" dataOnly="0" grandRow="1" axis="axisRow" fieldPosition="0" />

                    var appliesTo = XLPivotStyleFormatElement.All;
                    if (dataOnly)
                        appliesTo = XLPivotStyleFormatElement.Data;
                    else if (labelOnly)
                        appliesTo = XLPivotStyleFormatElement.Label;

                    var isRow = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.GrandRow, false);
                    var isColumn = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.GrandColumn, false);

                    // Either of the two should be true, else this is an unsupported format
                    if (!isRow && !isColumn)
                        continue;
                    //throw new NotImplementedException();

                    if (isRow)
                        styleFormat = pt.StyleFormats.RowGrandTotalFormats.ForElement(appliesTo) as XLPivotStyleFormat;
                    else
                        styleFormat = pt.StyleFormats.ColumnGrandTotalFormats.ForElement(appliesTo) as XLPivotStyleFormat;
                }
                else
                {
                    Int32 fieldIndex;
                    Boolean defaultSubtotal = false;

                    if (pivotArea.Field != null)
                        fieldIndex = (Int32)pivotArea.Field;
                    else if (pivotArea.PivotAreaReferences?.OfType<PivotAreaReference>()?.Any() ?? false)
                    {
                        // The field we want does NOT have any <x v="..."/>  children
                        var r = pivotArea.PivotAreaReferences.OfType<PivotAreaReference>().FirstOrDefault(r1 => !r1.Any());
                        if (r == null)
                            continue;

                        fieldIndex = Convert.ToInt32((UInt32)r.Field);
                        defaultSubtotal = OpenXmlHelper.GetBooleanValueAsBool(r.DefaultSubtotal, false);
                    }
                    else
                        throw new NotImplementedException();

                    XLPivotField field = null;
                    if (fieldIndex == -2)
                    {
                        var axis = pivotArea.Axis.Value;
                        if (axis == PivotTableAxisValues.AxisRow)
                            field = (XLPivotField)pt.RowLabels.Single(f => f.SourceName == "{{Values}}");
                        else if (axis == PivotTableAxisValues.AxisColumn)
                            field = (XLPivotField)pt.ColumnLabels.Single(f => f.SourceName == "{{Values}}");
                        else
                            continue;
                    }
                    else
                    {
                        if (fieldIndex >= pt.PivotCache.FieldCount)
                            throw PartStructureException.IncorrectAttributeValue();

                        var fieldName = pt.PivotCache.FieldNames[fieldIndex];
                        field = (XLPivotField)pt.ImplementedFields.SingleOrDefault(f => f.SourceName.Equals(fieldName));

                        if (field is null)
                            continue;
                    }

                    if (defaultSubtotal)
                    {
                        // Subtotal format
                        // Example:
                        // <x:pivotArea type="normal" fieldPosition="0">
                        //     <x:references count="1">
                        //         <x:reference field="0" defaultSubtotal="1" />
                        //     </x:references>
                        // </x:pivotArea>

                        styleFormat = field.StyleFormats.Subtotal as XLPivotStyleFormat;
                    }
                    else if (type == PivotAreaValues.Button)
                    {
                        // Header format
                        // Example:
                        // <x:pivotArea field="4" type="button" outline="0" axis="axisCol" fieldPosition="0" />
                        styleFormat = field.StyleFormats.Header as XLPivotStyleFormat;
                    }
                    else if (labelOnly)
                    {
                        // Label format
                        // Example:
                        // <x:pivotArea type="normal" dataOnly="0" labelOnly="1" fieldPosition="0">
                        //   <x:references count="1">
                        //     <x:reference field="4" />
                        //   </x:references>
                        // </x:pivotArea>
                        styleFormat = field.StyleFormats.Label as XLPivotStyleFormat;
                    }
                    else
                    {
                        // Assume DataValues format
                        // Example:
                        // <x:pivotArea type="normal" fieldPosition="0">
                        //     <x:references count="3">
                        //         <x:reference field="0" />
                        //         <x:reference field="4">
                        //             <x:x v="1" />
                        //         </x:reference>
                        //         <x:reference field="4294967294">
                        //             <x:x v="0" />
                        //         </x:reference>
                        //     </x:references>
                        //</x:pivotArea>
                        styleFormat = field.StyleFormats.DataValuesFormat as XLPivotStyleFormat;

                        foreach (var reference in pivotArea.PivotAreaReferences.OfType<PivotAreaReference>())
                        {
                            fieldIndex = unchecked((int)reference.Field.Value);
                            if (field.Offset == fieldIndex)
                                continue; // already handled

                            var fieldItem = reference.OfType<FieldItem>().First();
                            var fieldItemValue = fieldItem.Val.Value;

                            if (fieldIndex == -2)
                            {
                                // A value of -2 indicates the 'data' field.
                                styleFormat = (styleFormat as XLPivotValueStyleFormat)
                                    .ForValueField(pt.Values.ElementAt(checked((int)fieldItemValue)))
                                    as XLPivotValueStyleFormat;
                            }
                            else if (fieldIndex >= 0 && fieldIndex < pt.PivotCache.FieldCount)
                            {
                                var additionalFieldName = pt.PivotCache.FieldNames[fieldIndex];
                                var additionalField = pt.ImplementedFields
                                    .Single(f => f.SourceName == additionalFieldName);

                                Predicate<XLCellValue> predicate = null;
                                if (pt.PivotCache.TryGetFieldIndex(additionalFieldName, out var index))
                                {
                                    var values = pt.PivotCache.GetFieldSharedItems(index);
                                    if (fieldItemValue < values.Count)
                                    {
                                        predicate = o => o.ToString() == values[fieldItemValue].ToString();
                                    }
                                }

                                styleFormat = (styleFormat as XLPivotValueStyleFormat)
                                    .AndWith(additionalField, predicate)
                                    as XLPivotValueStyleFormat;
                            }
                            else
                            {
                                throw PartStructureException.IncorrectAttributeValue();
                            }
                        }
                    }

                    styleFormat.AreaType = type.Value.ToClosedXml();
                    styleFormat.Outline = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.Outline, true);
                    styleFormat.CollapsedLevelsAreSubtotals = OpenXmlHelper.GetBooleanValueAsBool(pivotArea.CollapsedLevelsAreSubtotals, false);
                }

                IXLStyle style = XLStyle.Default;
                if (format.FormatId != null)
                {
                    var df = differentialFormats[(Int32)format.FormatId.Value];
                    OpenXmlHelper.LoadFont(df.Font, style.Font);
                    OpenXmlHelper.LoadFill(df.Fill, style.Fill, differentialFillFormat: true);
                    OpenXmlHelper.LoadBorder(df.Border, style.Border);
                    OpenXmlHelper.LoadNumberFormat(df.NumberingFormat, style.NumberFormat);
                }

                styleFormat.Style = style;
            }
        }

        private static void LoadFieldOptions(PivotField pf, IXLPivotField pivotField)
        {
            if (pf.SubtotalCaption != null) pivotField.SubtotalCaption = pf.SubtotalCaption;
            if (pf.IncludeNewItemsInFilter != null) pivotField.IncludeNewItemsInFilter = pf.IncludeNewItemsInFilter.Value;
            if (pf.Outline != null) pivotField.Outline = pf.Outline.Value;
            if (pf.Compact != null) pivotField.Compact = pf.Compact.Value;
            if (pf.InsertBlankRow != null) pivotField.InsertBlankLines = pf.InsertBlankRow.Value;
            pivotField.ShowBlankItems = OpenXmlHelper.GetBooleanValueAsBool(pf.ShowAll, true);
            if (pf.InsertPageBreak != null) pivotField.InsertPageBreaks = pf.InsertPageBreak.Value;
            if (pf.SubtotalTop != null) pivotField.SubtotalsAtTop = pf.SubtotalTop.Value;
            if (pf.AllDrilled != null) pivotField.Collapsed = !pf.AllDrilled.Value;

            var pivotFieldExtensionList = pf.GetFirstChild<PivotFieldExtensionList>();
            var pivotFieldExtension = pivotFieldExtensionList?.GetFirstChild<PivotFieldExtension>();
            var field2010 = pivotFieldExtension?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotField>();
            if (field2010?.FillDownLabels != null) pivotField.RepeatItemLabels = field2010.FillDownLabels.Value;
        }

        private static void LoadSubtotals(PivotField pf, IXLPivotField pivotField)
        {
            if (pf.AverageSubTotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Average);
            if (pf.CountASubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Count);
            if (pf.CountSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.CountNumbers);
            if (pf.MaxSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Maximum);
            if (pf.MinSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Minimum);
            if (pf.ApplyStandardDeviationPInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.PopulationStandardDeviation);
            if (pf.ApplyVariancePInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.PopulationVariance);
            if (pf.ApplyProductInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Product);
            if (pf.ApplyStandardDeviationInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.StandardDeviation);
            if (pf.SumSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Sum);
            if (pf.ApplyVarianceInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Variance);

            if (pf.Items?.Any() ?? false)
            {
                var items = pf.Items.OfType<Item>().Where(i => i.Index != null && i.Index.HasValue);
                if (!items.Any(i => i.HideDetails == null || BooleanValue.ToBoolean(i.HideDetails)))
                    pivotField.SetCollapsed();
            }
        }

        private static void ReadCacheFields(CacheFields cacheFields, XLPivotCache pivotCache)
        {
            foreach (var cacheField in cacheFields.Elements<CacheField>())
            {
                if (cacheField.Name?.Value is not { } fieldName)
                    throw PartStructureException.MissingAttribute();

                if (pivotCache.ContainsField(fieldName))
                {
                    // We don't allow duplicate field names... but what do we do if we find one? Let's just skip it.
                    continue;
                }

                var fieldStats = ReadCacheFieldStats(cacheField);
                var fieldSharedItems = cacheField.SharedItems is not null
                    ? ReadSharedItems(cacheField)
                    : new XLPivotCacheSharedItems();

                var fieldValues = new XLPivotCacheValues(fieldSharedItems, fieldStats);
                pivotCache.AddCachedField(fieldName, fieldValues);
            }
        }

        private static XLPivotCacheValuesStats ReadCacheFieldStats(CacheField cacheField)
        {
            var sharedItems = cacheField.SharedItems;

            // Various statistics about the records of the field, not just shared items.
            var containsBlank = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsBlank, false);
            var containsNumber = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsNumber, false);
            var containsOnlyInteger = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsInteger, false);
            var minValue = sharedItems?.MinValue?.Value;
            var maxValue = sharedItems?.MaxValue?.Value;
            var containsDate = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsDate, false);
            var minDate = sharedItems?.MinDate?.Value;
            var maxDate = sharedItems?.MaxDate?.Value;
            var containsString = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.ContainsString, true);
            var longText = OpenXmlHelper.GetBooleanValueAsBool(sharedItems?.LongText, false);

            // The containsMixedTypes, containsNonDate and containsSemiMixedTypes are derived from primary stats.
            return new XLPivotCacheValuesStats(
                containsBlank,
                containsNumber,
                containsOnlyInteger,
                minValue,
                maxValue,
                containsString,
                longText,
                containsDate,
                minDate,
                maxDate);
        }

        private static XLPivotCacheSharedItems ReadSharedItems(CacheField cacheField)
        {
            var sharedItems = new XLPivotCacheSharedItems();

            // If there are no shared items, the cache record can't contain field items
            // referencing the shared items.
            if (cacheField.SharedItems is not { } fieldSharedItems)
                return sharedItems;

            foreach (var item in fieldSharedItems.Elements())
            {
                // Shared items can't contain element of type index (`x`),
                // because index references shared items. That is main reason
                // for rather significant duplication with reading records.
                switch (item)
                {
                    case MissingItem:
                        sharedItems.AddMissing();
                        break;

                    case NumberItem numberItem:
                        if (numberItem.Val?.Value is not { } number)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddNumber(number);
                        break;

                    case BooleanItem booleanItem:
                        if (booleanItem.Val?.Value is not { } boolean)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddBoolean(boolean);
                        break;

                    case ErrorItem errorItem:
                        if (errorItem.Val?.Value is not { } errorText)
                            throw PartStructureException.MissingAttribute();

                        if (!XLErrorParser.TryParseError(errorText, out var error))
                            throw PartStructureException.IncorrectAttributeFormat();

                        sharedItems.AddError(error);
                        break;

                    case StringItem stringItem:
                        if (stringItem.Val?.Value is not { } text)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddString(text);
                        break;

                    case DateTimeItem dateTimeItem:
                        if (dateTimeItem.Val?.Value is not { } dateTime)
                            throw PartStructureException.MissingAttribute();

                        sharedItems.AddDateTime(dateTime);
                        break;

                    default:
                        throw PartStructureException.ExpectedElementNotFound();
                }
            }

            return sharedItems;
        }

        private static void ReadRecords(PivotCacheRecords recordsPart, XLPivotCache pivotCache)
        {
            // Number of records can be rather large, preallocate capacity to avoid reallocation.
            var recordCount = recordsPart.Count?.Value is not null
                ? checked((int)recordsPart.Count.Value)
                : 0;
            pivotCache.AllocateRecordCapacity(recordCount);

            var fieldsCount = pivotCache.FieldCount;
            foreach (var record in recordsPart.Elements<PivotCacheRecord>())
            {
                var recordColumns = record.ChildElements.Count;
                if (recordColumns != fieldsCount)
                    throw PartStructureException.IncorrectElementsCount();

                for (var fieldIdx = 0; fieldIdx < fieldsCount; ++fieldIdx)
                {
                    var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                    var recordItem = record.ElementAt(fieldIdx);

                    switch (recordItem)
                    {
                        case MissingItem:
                            fieldValues.AddMissing();
                            break;

                        case NumberItem numberItem:
                            if (numberItem.Val?.Value is not { } number)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddNumber(number);
                            break;

                        case BooleanItem booleanItem:
                            if (booleanItem.Val?.Value is not { } boolean)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddBoolean(boolean);
                            break;

                        case ErrorItem errorItem:
                            if (errorItem.Val?.Value is not { } errorText)
                                throw PartStructureException.MissingAttribute();

                            if (!XLErrorParser.TryParseError(errorText, out var error))
                                throw PartStructureException.IncorrectAttributeFormat();

                            fieldValues.AddError(error);
                            break;

                        case StringItem stringItem:
                            if (stringItem.Val?.Value is not { } text)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddString(text);
                            break;

                        case DateTimeItem dateTimeItem:
                            if (dateTimeItem.Val?.Value is not { } dateTime)
                                throw PartStructureException.MissingAttribute();

                            fieldValues.AddDateTime(dateTime);
                            break;

                        case FieldItem indexItem:
                            if (indexItem.Val?.Value is not { } index)
                                throw PartStructureException.MissingAttribute();

                            if (index >= fieldValues.SharedCount)
                                throw PartStructureException.IncorrectAttributeValue();

                            fieldValues.AddIndex(index);
                            break;

                        default:
                            throw PartStructureException.ExpectedElementNotFound();
                    }
                }
            }
        }
    }
}
