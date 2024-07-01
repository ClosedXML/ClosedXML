#nullable disable
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.IO;

internal class PivotTableDefinitionPartReader
{
    /// <summary>
    /// A field displayed as <c>∑Values</c> in a pivot table that contains names of all aggregation
    /// function in value fields collection. Also commonly called 'data' field.
    /// </summary>
    private const int ValuesFieldIndex = -2;

    internal static void Load(WorkbookPart workbookPart, Dictionary<int, DifferentialFormat> differentialFormats, PivotTablePart pivotTablePart, WorksheetPart worksheetPart, XLWorksheet ws, LoadContext context)
    {
        var workbook = ws.Workbook;
        var cache = pivotTablePart.PivotTableCacheDefinitionPart;
        var cacheDefinitionRelId = workbookPart.GetIdOfPart(cache);

        var pivotSource = workbook.PivotCachesInternal
            .FirstOrDefault<XLPivotCache>(ps => ps.WorkbookCacheRelId == cacheDefinitionRelId);

        if (pivotSource == null)
        {
            // If it's missing, find a 'similar' pivot cache, i.e. one that's based on the same source range/table
            pivotSource = workbook.PivotCachesInternal
                .FirstOrDefault<XLPivotCache>(ps =>
                    ps.PivotSourceReference.Equals(PivotTableCacheDefinitionPartReader.ParsePivotSourceReference(cache)));
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
            var pt = LoadPivotTableDefinition(pivotTableDefinition, ws, pivotSource, differentialFormats, context);
            ws.PivotTables.Add(pt);

            pt.RelId = worksheetPart.GetIdOfPart(pivotTablePart);
            pt.CacheDefinitionRelId = pivotTablePart.GetIdOfPart(cache);
        }
    }
/*
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
                                        pivotField.SetSort(pf.SortType.Value.ToClosedXml());
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
                                    pivotField.SetSort(pf.SortType.Value.ToClosedXml());
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
*/
#nullable enable
    private static XLPivotTable LoadPivotTableDefinition(PivotTableDefinition pivotTable, XLWorksheet sheet, XLPivotCache cache, Dictionary<int, DifferentialFormat> differentialFormats, LoadContext context)
    {
        // Load base attributes
        var xlPivotTable = LoadPivotTableAttributes(pivotTable, sheet, cache);

        // Load location
        var location = pivotTable.Location;
        if (location is null)
            throw PartStructureException.ExpectedElementNotFound();

        var referenceText = location.Reference?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.Area = XLSheetRange.Parse(referenceText);
        xlPivotTable.FirstHeaderRow = location.FirstHeaderRow?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataRow = location.FirstDataRow?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataCol = location.FirstDataColumn?.Value ?? throw PartStructureException.MissingAttribute();

        // Skip `rowPageCount` and `colPageCount`, because they are derived from filterAreaOrder, filterFieldsPageWrap and pageField count

        // Load pivot fields
        var pivotFields = pivotTable.PivotFields;
        if (pivotFields is not null)
        {
            foreach (var pivotField in pivotFields.Cast<PivotField>())
                xlPivotTable.AddField(LoadPivotField(pivotField, xlPivotTable));
        }

        // Load row axis fields and items
        LoadAxisFields(pivotTable.RowFields, xlPivotTable.RowAxis, xlPivotTable);
        LoadAxisItems(pivotTable.RowItems, xlPivotTable.RowAxis);

        // Load column axis fields and items
        LoadAxisFields(pivotTable.ColumnFields, xlPivotTable.ColumnAxis, xlPivotTable);
        LoadAxisItems(pivotTable.ColumnItems, xlPivotTable.ColumnAxis);

        // Load page fields, i.e. the filters region.
        var pageFields = pivotTable.PageFields;
        if (pageFields is not null)
        {
            foreach (var pageField in pageFields.Cast<PageField>())
            {
                var field = pageField.Field?.Value ?? throw PartStructureException.MissingAttribute();
                var itemIndex = pageField.Item?.Value;
                var hierarchyIndex = pageField.Hierarchy?.Value;
                var hierarchyUniqueName = pageField.Name;
                var hierarchyDisplayName = pageField.Caption;
                var xlPageField = new XLPivotPageField(field)
                {
                    ItemIndex = itemIndex,
                    HierarchyIndex = hierarchyIndex,
                    HierarchyUniqueName = hierarchyUniqueName,
                    HierarchyDisplayName = hierarchyDisplayName,
                };
                xlPivotTable.Filters.AddField(xlPageField);
            }
        }

        // Load data fields.
        var dataFields = pivotTable.DataFields;
        if (dataFields is not null)
        {
            foreach (var dataField in dataFields.Cast<DataField>())
            {
                var name = dataField.Name?.Value;
                var field = dataField.Field?.Value ?? throw PartStructureException.MissingAttribute();
                var subtotal = dataField.Subtotal?.Value.ToClosedXml() ?? XLPivotSummary.Sum;
                var showDataAsFormat = dataField.ShowDataAs?.Value.ToClosedXml() ?? XLPivotCalculation.Normal;
                var baseField = dataField.BaseField?.Value ?? -1;
                var baseItem = dataField.BaseItem?.Value ?? 1048832;
                var numberFormatId = dataField.NumberFormatId?.Value;
                var numberFormat = context.GetNumberFormat(numberFormatId);
                var xlDataField = new XLPivotDataField(xlPivotTable, checked((int)field))
                {
                    DataFieldName = name,
                    Subtotal = subtotal,
                    ShowDataAsFormat = showDataAsFormat,
                    BaseField = baseField,
                    BaseItem = baseItem,
                    NumberFormatValue = numberFormat,
                };
                xlPivotTable.DataFields.AddField(xlDataField);
            }
        }

        // Load formats
        var formats = pivotTable.Formats;
        if (formats is not null)
        {
            foreach (var format in formats.Cast<Format>())
            {
                var action = format.Action?.Value.ToClosedXml() ?? XLPivotFormatAction.Formatting;
                var dxfStyle = XLStyle.Default;
                if (format.FormatId is not null)
                {
                    // TODO: What about alignment?
                    var df = differentialFormats[checked((int)format.FormatId.Value)];
                    OpenXmlHelper.LoadFont(df.Font, dxfStyle.Font);
                    OpenXmlHelper.LoadFill(df.Fill, dxfStyle.Fill, differentialFillFormat: true);
                    OpenXmlHelper.LoadBorder(df.Border, dxfStyle.Border);
                    OpenXmlHelper.LoadNumberFormat(df.NumberingFormat, dxfStyle.NumberFormat);
                }

                var pivotArea = format.PivotArea ?? throw PartStructureException.ExpectedElementNotFound();
                var xlPivotArea = LoadPivotArea(pivotArea);
                var xlFormat = new XLPivotFormat(xlPivotArea)
                {
                    Action = action,
                    DxfStyle = dxfStyle,
                };
                xlPivotTable.AddFormat(xlFormat);
            }
        }

        var conditionalFormats = pivotTable.ConditionalFormats;
        if (conditionalFormats is not null)
        {
            foreach (var conditionalFormat in conditionalFormats.Cast<ConditionalFormat>())
            {
                var scope = conditionalFormat.Scope?.Value.ToClosedXml() ?? XLPivotCfScope.SelectedCells;
                var type = conditionalFormat.Type?.Value.ToClosedXml() ?? XLPivotCfRuleType.None;
                var priority = conditionalFormat.Priority?.Value ?? throw PartStructureException.MissingAttribute();
                var format = context.GetPivotCf(sheet.Name, checked((int)priority));
                var xlConditionalFormat = new XLPivotConditionalFormat(format)
                {
                    Scope = scope,
                    Type = type,
                };
                var pivotAreas = conditionalFormat.PivotAreas;
                if (pivotAreas is not null)
                {
                    foreach (var pivotArea in pivotAreas.Cast<PivotArea>())
                    {
                        var xlPivotArea = LoadPivotArea(pivotArea);
                        xlConditionalFormat.AddArea(xlPivotArea);
                    }
                }

                xlPivotTable.AddConditionalFormat(xlConditionalFormat);
            }
        }

        // TODO: chartFormats
        // pivotHierarchies is OLAP and thus for now out of scope.
        var pivotTableStyle = pivotTable.GetFirstChild<PivotTableStyle>();
        LoadPivotTableStyle(pivotTableStyle, xlPivotTable);

        // TODO: filters
        // rowHierarchiesUsage is OLAP and thus for now out of scope.
        // colHierarchiesUsage is OLAP and thus for now out of scope.
        LoadExtensionList(pivotTable, xlPivotTable);

        return xlPivotTable;
    }

    private static XLPivotTable LoadPivotTableAttributes(PivotTableDefinition pivotTable, XLWorksheet sheet, XLPivotCache cache)
    {
        var name = pivotTable.Name?.Value ?? throw PartStructureException.MissingAttribute();
        var cacheId = pivotTable.CacheId?.Value ?? throw PartStructureException.MissingAttribute();
        var dataOnRows = pivotTable.DataOnRows?.Value ?? false;

        // DataPosition attribute is skipped, because it basically represents a field on one of axis.
        // Excel requires that dataPosition and field with index -2 must be in list of respective axis
        // at correct place, otherwise it crashes. To make things simple, we set the value when it is
        // encountered on the correct axis (plus there is a check that field is not used on multiple axes
        // that would cause exception).
        var autoFormatId = pivotTable.AutoFormatId?.Value;
        var applyNumberFormats = pivotTable.ApplyNumberFormats?.Value ?? false;
        var applyBorderFormats = pivotTable.ApplyBorderFormats?.Value ?? false;
        var applyFontFormats = pivotTable.ApplyFontFormats?.Value ?? false;
        var applyPatternFormats = pivotTable.ApplyPatternFormats?.Value ?? false;
        var applyAlignmentFormats = pivotTable.ApplyAlignmentFormats?.Value ?? false;
        var applyWidthHeightFormats = pivotTable.ApplyWidthHeightFormats?.Value ?? false;
        var dataCaption = pivotTable.DataCaption?.Value ?? throw PartStructureException.MissingAttribute();
        var grandTotalCaption = pivotTable.GrandTotalCaption?.Value;
        var errorCaption = pivotTable.ErrorCaption?.Value;
        var showError = pivotTable.ShowError?.Value ?? false;
        var missingCaption = pivotTable.MissingCaption?.Value ?? string.Empty;
        var showMissing = pivotTable.ShowMissing?.Value ?? true;
        var pageStyle = pivotTable.PageStyle?.Value;
        var pivotTableStyleName = pivotTable.PivotTableStyleName?.Value;
        var vacatedStyle = pivotTable.VacatedStyle?.Value;
        var tag = pivotTable.Tag?.Value;
        var updatedVersion = pivotTable.UpdatedVersion?.Value ?? 0;
        var minRefreshableVersion = pivotTable.MinRefreshableVersion?.Value ?? 0;
        var asteriskTotals = pivotTable.AsteriskTotals?.Value ?? false;
        var showItems = pivotTable.ShowItems?.Value ?? true;
        var editData = pivotTable.EditData?.Value ?? false;
        var disableFieldList = pivotTable.DisableFieldList?.Value ?? false;
        var showCalculatedMembers = pivotTable.ShowCalculatedMembers?.Value ?? true;
        var visualTotals = pivotTable.VisualTotals?.Value ?? true;
        var showMultipleLabel = pivotTable.ShowMultipleLabel?.Value ?? true;
        var showDataDropDown = pivotTable.ShowDataDropDown?.Value ?? true;
        var showDrill = pivotTable.ShowDrill?.Value ?? true;
        var printDrill = pivotTable.PrintDrill?.Value ?? false;
        var showMemberPropertyTips = pivotTable.ShowMemberPropertyTips?.Value ?? true;
        var showDataTips = pivotTable.ShowDataTips?.Value ?? true;
        var enableWizard = pivotTable.EnableWizard?.Value ?? true;
        var enableDrill = pivotTable.EnableDrill?.Value ?? true;
        var enableFieldProperties = pivotTable.EnableFieldProperties?.Value ?? true;
        var preserveFormatting = pivotTable.PreserveFormatting?.Value ?? true;
        var useAutoFormatting = pivotTable.UseAutoFormatting?.Value ?? false;
        var pageWrap = pivotTable.PageWrap?.Value ?? 0;
        var pageOverThenDown = pivotTable.PageOverThenDown?.Value ?? false;
        var subtotalHiddenItems = pivotTable.SubtotalHiddenItems?.Value ?? false;
        var rowGrandTotals = pivotTable.RowGrandTotals?.Value ?? true;
        var columnGrandTotals = pivotTable.ColumnGrandTotals?.Value ?? true;
        var fieldPrintTitles = pivotTable.FieldPrintTitles?.Value ?? false;
        var itemPrintTitles = pivotTable.ItemPrintTitles?.Value ?? false;
        var mergeItem = pivotTable.MergeItem?.Value ?? false;
        var showDropZones = pivotTable.ShowDropZones?.Value ?? true;
        var createdVersion = pivotTable.CreatedVersion?.Value ?? 0;
        var indent = pivotTable.Indent?.Value ?? 1;
        var showEmptyRow = pivotTable.ShowEmptyRow?.Value ?? false;
        var showEmptyColumn = pivotTable.ShowEmptyColumn?.Value ?? false;
        var showHeaders = pivotTable.ShowHeaders?.Value ?? true;
        var compact = pivotTable.Compact?.Value ?? true;
        var outline = pivotTable.Outline?.Value ?? false;
        var outlineData = pivotTable.OutlineData?.Value ?? false;
        var compactData = pivotTable.CompactData?.Value ?? true;
        var published = pivotTable.Published?.Value ?? false;
        var gridDropZones = pivotTable.GridDropZones?.Value ?? false;
        var stopImmersiveUi = pivotTable.StopImmersiveUi?.Value ?? true;
        var multipleFieldFilters = pivotTable.MultipleFieldFilters?.Value ?? true;
        var chartFormat = pivotTable.ChartFormat?.Value ?? 0;
        var rowHeaderCaption = pivotTable.RowHeaderCaption?.Value;
        var columnHeaderCaption = pivotTable.ColumnHeaderCaption?.Value;
        var fieldListSortAscending = pivotTable.FieldListSortAscending?.Value ?? false;
        var mdxSubQueries = pivotTable.MdxSubqueries?.Value ?? false;
        var customSortList = pivotTable.CustomListSort?.Value ?? true;

        var xlPivotTable = new XLPivotTable(sheet, cache)
        {
            Name = name,
            DataOnRows = dataOnRows,
            DataPosition = null, // 'data' field is set when during axis loading (if present).
            AutoFormatId = autoFormatId,
            ApplyNumberFormats = applyNumberFormats,
            ApplyBorderFormats = applyBorderFormats,
            ApplyFontFormats = applyFontFormats,
            ApplyPatternFormats = applyPatternFormats,
            ApplyAlignmentFormats = applyAlignmentFormats,
            ApplyWidthHeightFormats = applyWidthHeightFormats,
            DataCaption = dataCaption,
            GrandTotalCaption = grandTotalCaption,
            ErrorValueReplacement = errorCaption,
            ShowError = showError,
            MissingCaption = missingCaption,
            ShowMissing = showMissing,
            PageStyle = pageStyle,
            PivotTableStyleName = pivotTableStyleName,
            VacatedStyle = vacatedStyle,
            Tag = tag,
            UpdatedVersion = updatedVersion,
            MinRefreshableVersion = minRefreshableVersion,
            AsteriskTotals = asteriskTotals,
            DisplayItemLabels = showItems,
            EditData = editData,
            DisableFieldList = disableFieldList,
            ShowCalculatedMembers = showCalculatedMembers,
            VisualTotals = visualTotals,
            ShowMultipleLabel = showMultipleLabel,
            ShowDataDropDown = showDataDropDown,
            ShowExpandCollapseButtons = showDrill,
            PrintExpandCollapsedButtons = printDrill,
            ShowPropertiesInTooltips = showMemberPropertyTips,
            ShowContextualTooltips = showDataTips,
            EnableEditingMechanism = enableWizard,
            EnableShowDetails = enableDrill,
            EnableFieldProperties = enableFieldProperties,
            PreserveCellFormatting = preserveFormatting,
            AutofitColumns = useAutoFormatting,
            FilterFieldsPageWrap = checked((int)pageWrap),
            FilterAreaOrder = pageOverThenDown ? XLFilterAreaOrder.OverThenDown : XLFilterAreaOrder.DownThenOver,
            FilteredItemsInSubtotals = subtotalHiddenItems,
            ShowGrandTotalsRows = rowGrandTotals,
            ShowGrandTotalsColumns = columnGrandTotals,
            PrintTitles = fieldPrintTitles,
            RepeatRowLabels = itemPrintTitles,
            MergeAndCenterWithLabels = mergeItem,
            ShowDropZones = showDropZones,
            PivotCacheCreatedVersion = createdVersion,
            RowLabelIndent = checked((int)indent),
            ShowEmptyItemsOnRows = showEmptyRow,
            ShowEmptyItemsOnColumns = showEmptyColumn,
            DisplayCaptionsAndDropdowns = showHeaders,
            Compact = compact,
            Outline = outline,
            OutlineData = outlineData,
            CompactData = compactData,
            Published = published,
            ClassicPivotTableLayout = gridDropZones,
            StopImmersiveUi = stopImmersiveUi,
            AllowMultipleFilters = multipleFieldFilters,
            ChartFormat = chartFormat,
            RowHeaderCaption = rowHeaderCaption,
            ColumnHeaderCaption = columnHeaderCaption,
            SortFieldsAtoZ = fieldListSortAscending,
            MdxSubQueries = mdxSubQueries,
            UseCustomListsForSorting = customSortList,
        };
        return xlPivotTable;
    }

    private static XLPivotTableField LoadPivotField(PivotField pivotField, XLPivotTable xlPivotTable)
    {
        var customName = pivotField.Name?.Value;
        var axis = pivotField.Axis?.Value.ToClosedXml();
        var dataField = pivotField.DataField?.Value ?? false;
        var subtotalCaption = pivotField.SubtotalCaption?.Value;
        var showDropDowns = pivotField.ShowDropDowns?.Value ?? true;
        var hiddenLevel = pivotField.HiddenLevel?.Value ?? false;
        var uniqueMemberProperty = pivotField.UniqueMemberProperty?.Value;
        var compact = pivotField.Compact?.Value ?? true;
        var allDrilled = pivotField.AllDrilled?.Value ?? false;
        var numberFormatId = pivotField.NumberFormatId?.Value; // TODO: Load properly number format
        var outline = pivotField.Outline?.Value ?? true;
        var subtotalTop = pivotField.SubtotalTop?.Value ?? true;
        var dragToRow = pivotField.DragToRow?.Value ?? true;
        var dragToColumn = pivotField.DragToColumn?.Value ?? true;
        var multipleItemSelectionAllowed = pivotField.MultipleItemSelectionAllowed?.Value ?? false;
        var dragToPage = pivotField.DragToPage?.Value ?? true;
        var dragToData = pivotField.DragToData?.Value ?? true;
        var dragOff = pivotField.DragOff?.Value ?? true;
        var showAll = pivotField.ShowAll?.Value ?? true;
        var insertBlankRow = pivotField.InsertBlankRow?.Value ?? false;
        var serverField = pivotField.ServerField?.Value ?? false;
        var insertPageBreak = pivotField.InsertPageBreak?.Value ?? false;
        var autoShow = pivotField.AutoShow?.Value ?? false;
        var topAutoShow = pivotField.TopAutoShow?.Value ?? true;
        var hideNewItems = pivotField.HideNewItems?.Value ?? false;
        var measureFilter = pivotField.MeasureFilter?.Value ?? false;
        var includeNewItemsInFilter = pivotField.IncludeNewItemsInFilter?.Value ?? false;
        var itemPageCount = pivotField.ItemPageCount?.Value ?? 10u;
        var sortType = pivotField.SortType?.Value.ToClosedXml() ?? XLPivotSortType.Default;
        var dataSourceSort = pivotField.DataSourceSort?.Value;
        var nonAutoSortDefault = pivotField.NonAutoSortDefault?.Value ?? false;
        var rankBy = pivotField.RankBy?.Value;
        var defaultSubtotal = pivotField.DefaultSubtotal?.Value ?? true;
        var sumSubtotal = pivotField.SumSubtotal?.Value ?? false;
        var countASubtotal = pivotField.CountASubtotal?.Value ?? false;
        var avgSubtotal = pivotField.AverageSubTotal?.Value ?? false;
        var maxSubtotal = pivotField.MaxSubtotal?.Value ?? false;
        var minSubtotal = pivotField.MinSubtotal?.Value ?? false;
        var productSubtotal = pivotField.ApplyProductInSubtotal?.Value ?? false;
        var countSubtotal = pivotField.CountSubtotal?.Value ?? false;
        var stdDevSubtotal = pivotField.ApplyStandardDeviationInSubtotal?.Value ?? false;
        var stdDevPSubtotal = pivotField.ApplyStandardDeviationPInSubtotal?.Value ?? false;
        var varSubtotal = pivotField.ApplyVarianceInSubtotal?.Value ?? false;
        var varPSubtotal = pivotField.ApplyVariancePInSubtotal?.Value ?? false;
        var showPropCell = pivotField.ShowPropCell?.Value ?? false;
        var showPropTip = pivotField.ShowPropertyTooltip?.Value ?? false;
        var showPropAsCaption = pivotField.ShowPropAsCaption?.Value ?? false;
        var defaultAttributeDrillState = pivotField.DefaultAttributeDrillState?.Value ?? false;

        var subtotals = new HashSet<XLSubtotalFunction>();
        if (defaultSubtotal)
            subtotals.Add(XLSubtotalFunction.Automatic);

        if (sumSubtotal)
            subtotals.Add(XLSubtotalFunction.Sum);

        if (countASubtotal)
            subtotals.Add(XLSubtotalFunction.Count);

        if (avgSubtotal)
            subtotals.Add(XLSubtotalFunction.Average);

        if (maxSubtotal)
            subtotals.Add(XLSubtotalFunction.Maximum);

        if (minSubtotal)
            subtotals.Add(XLSubtotalFunction.Minimum);

        if (productSubtotal)
            subtotals.Add(XLSubtotalFunction.Product);

        if (countSubtotal)
            subtotals.Add(XLSubtotalFunction.CountNumbers);

        if (stdDevSubtotal)
            subtotals.Add(XLSubtotalFunction.StandardDeviation);

        if (stdDevPSubtotal)
            subtotals.Add(XLSubtotalFunction.PopulationStandardDeviation);

        if (varSubtotal)
            subtotals.Add(XLSubtotalFunction.Variance);

        if (varPSubtotal)
            subtotals.Add(XLSubtotalFunction.PopulationVariance);

        var xlField = new XLPivotTableField(xlPivotTable)
        {
            Name = customName,
            Axis = axis,
            DataField = dataField,
            SubtotalCaption = subtotalCaption ?? string.Empty,
            ShowDropDowns = showDropDowns,
            HiddenLevel = hiddenLevel,
            UniqueMemberProperty = uniqueMemberProperty,
            Compact = compact,
            AllDrilled = allDrilled,
            NumberFormatId = numberFormatId,
            Outline = outline,
            SubtotalTop = subtotalTop,
            DragToRow = dragToRow,
            DragToColumn = dragToColumn,
            MultipleItemSelectionAllowed = multipleItemSelectionAllowed,
            DragToPage = dragToPage,
            DragToData = dragToData,
            DragOff = dragOff,
            ShowAll = showAll,
            InsertBlankRow = insertBlankRow,
            ServerField = serverField,
            InsertPageBreak = insertPageBreak,
            AutoShow = autoShow,
            TopAutoShow = topAutoShow,
            HideNewItems = hideNewItems,
            MeasureFilter = measureFilter,
            IncludeNewItemsInFilter = includeNewItemsInFilter,
            ItemPageCount = itemPageCount,
            SortType = sortType,
            DataSourceSort = dataSourceSort,
            NonAutoSortDefault = nonAutoSortDefault,
            RankBy = rankBy,
            Subtotals = subtotals,
            ShowPropCell = showPropCell,
            ShowPropTip = showPropTip,
            ShowPropAsCaption = showPropAsCaption,
            DefaultAttributeDrillState = defaultAttributeDrillState,
        };

        var items = pivotField.Items;
        if (items is not null)
        {
            foreach (var item in items.Cast<Item>())
            {
                // Attributes `sd` and `d` were swapped in spec.
                var approximatelyHasChildren = item.ChildItems?.Value ?? false;
                var details = item.Expanded?.Value ?? false;
                var drillAcrossAttributes = item.DrillAcrossAttributes?.Value ?? true;
                var calculatedMember = item.Calculated?.Value ?? false;
                var hidden = item.Hidden?.Value ?? false;
                var missing = item.Missing?.Value ?? false;
                var itemUserCaption = item.ItemName;
                var valueIsString = item.HasStringVlue?.Value ?? false;
                var showDetails = item.HideDetails?.Value ?? true;
                var itemIndex = item.Index?.Value;
                var itemType = item.ItemType?.Value.ToClosedXml() ?? XLPivotItemType.Data;
                var xlItem = new XLPivotFieldItem(xlField, itemIndex is null ? null : checked((int)itemIndex.Value))
                {
                    ApproximatelyHasChildren = approximatelyHasChildren,
                    Details = details,
                    DrillAcrossAttributes = drillAcrossAttributes,
                    CalculatedMember = calculatedMember,
                    Hidden = hidden,
                    Missing = missing,
                    ItemUserCaption = itemUserCaption,
                    ValueIsString = valueIsString,
                    ShowDetails = showDetails,
                    ItemType = itemType,
                };

                xlField.AddItem(xlItem);
            }
        }

        // TODO: autoSortScope

        // extLst
        var pivotFieldExtensionList = pivotField.GetFirstChild<PivotFieldExtensionList>();
        var pivotFieldExtension = pivotFieldExtensionList?.GetFirstChild<PivotFieldExtension>();
        var field2010 = pivotFieldExtension?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotField>();
        xlField.RepeatItemLabels = field2010?.FillDownLabels?.Value ?? false;

        return xlField;
    }

    private static void LoadAxisFields(OpenXmlCompositeElement? fields, XLPivotTableAxis axis, XLPivotTable xlPivotTable)
    {
        if (fields is not null)
        {
            foreach (var field in fields.Cast<Field>())
            {
                // Axis can contain 'data' field.
                var fieldIndex = field.Index?.Value ?? throw PartStructureException.MissingAttribute();
                if (fieldIndex >= xlPivotTable.PivotFields.Count || (fieldIndex < 0 && fieldIndex != ValuesFieldIndex))
                    throw PartStructureException.IncorrectAttributeValue();

                axis.AddField(fieldIndex);
            }
        }
    }

    private static void LoadAxisItems(OpenXmlCompositeElement? axisItems, XLPivotTableAxis axis)
    {
        if (axisItems is not null)
        {
            // Both row and column use RowItem type for axis item.
            var previous = new List<int>();
            foreach (var axisItem in axisItems.Cast<RowItem>())
            {
                var xlItemType = axisItem.ItemType?.Value.ToClosedXml() ?? XLPivotItemType.Data;
                var dataFieldIndex = axisItem.Index?.Value ?? 0; // This is used by 'data' field
                var repeatedCount = axisItem.RepeatedItemCount?.Value ?? 0;
                var fieldIndexes = new List<int>();
                foreach (var dataIndex in axisItem.ChildElements.Cast<MemberPropertyIndex>())
                    fieldIndexes.Add(dataIndex.Val?.Value ?? 0);

                var allFieldIndexes = previous.Take((int)repeatedCount).Concat(fieldIndexes).ToList();
                axis.AddItem(new XLPivotFieldAxisItem(xlItemType, dataFieldIndex, allFieldIndexes));
                previous = allFieldIndexes;
            }
        }
    }

    private static XLPivotArea LoadPivotArea(PivotArea pivotArea)
    {
        var field = pivotArea.Field?.Value;
        var type = pivotArea.Type?.Value.ToClosedXml() ?? XLPivotAreaType.Normal;
        var dataOnly = pivotArea.DataOnly?.Value ?? true;
        var labelOnly = pivotArea.LabelOnly?.Value ?? false;
        var grandRow = pivotArea.GrandRow?.Value ?? false;
        var grandCol = pivotArea.GrandColumn?.Value ?? false;
        var cacheIndex = pivotArea.CacheIndex?.Value ?? false;
        var outline = pivotArea.Outline?.Value ?? true;
        var offset = pivotArea.Offset?.Value is { } offsetRefText ? XLSheetRange.Parse(offsetRefText) : (XLSheetRange?)null;
        var collapsedLevelsAreSubtotals = pivotArea.CollapsedLevelsAreSubtotals?.Value ?? false;
        var axis = pivotArea.Axis?.Value.ToClosedXml();
        var fieldPosition = pivotArea.FieldPosition?.Value;
        var xlPivotArea = new XLPivotArea
        {
            Field = field,
            Type = type,
            DataOnly = dataOnly,
            LabelOnly = labelOnly,
            GrandRow = grandRow,
            GrandCol = grandCol,
            CacheIndex = cacheIndex,
            Outline = outline,
            Offset = offset,
            CollapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals,
            Axis = axis,
            FieldPosition = fieldPosition
        };

        // Can contain extensions, in theory at least.
        var references = pivotArea.PivotAreaReferences;
        if (references is not null)
        {
            foreach (var reference in references.Cast<PivotAreaReference>())
                xlPivotArea.AddReference(LoadPivotReference(reference));
        }

        return xlPivotArea;
    }

    private static XLPivotReference LoadPivotReference(PivotAreaReference reference)
    {
        var field = reference.Field?.Value;
        var selected = reference.Selected?.Value ?? true;
        var byPosition = reference.ByPosition?.Value ?? false;
        var relative = reference.Relative?.Value ?? false;
        var defaultSubtotal = reference.DefaultSubtotal?.Value ?? false;
        var sumSubtotal = reference.SumSubtotal?.Value ?? false;
        var countASubtotal = reference.CountASubtotal?.Value ?? false;
        var avgSubtotal = reference.AverageSubtotal?.Value ?? false;
        var maxSubtotal = reference.MaxSubtotal?.Value ?? false;
        var minSubtotal = reference.MinSubtotal?.Value ?? false;
        var productSubtotal = reference.ApplyProductInSubtotal?.Value ?? false;
        var countSubtotal = reference.CountSubtotal?.Value ?? false;
        var stdDevSubtotal = reference.ApplyStandardDeviationInSubtotal?.Value ?? false;
        var stdDevPSubtotal = reference.ApplyStandardDeviationPInSubtotal?.Value ?? false;
        var varSubtotal = reference.ApplyVarianceInSubtotal?.Value ?? false;
        var varPSubtotal = reference.ApplyVariancePInSubtotal?.Value ?? false;

        var xlReference = new XLPivotReference
        {
            Field = field,
            Selected = selected,
            ByPosition = byPosition,
            Relative = relative,
            DefaultSubtotal = defaultSubtotal,
            SumSubtotal = sumSubtotal,
            CountASubtotal = countASubtotal,
            AvgSubtotal = avgSubtotal,
            MaxSubtotal = maxSubtotal,
            MinSubtotal = minSubtotal,
            ProductSubtotal = productSubtotal,
            CountSubtotal = countSubtotal,
            StdDevSubtotal = stdDevSubtotal,
            StdDevPSubtotal = stdDevPSubtotal,
            VarSubtotal = varSubtotal,
            VarPSubtotal = varPSubtotal,
        };

        // Add indexes after the reference is initialized, so it can check values by cacheIndex/byPosition.
        foreach (var fieldItem in reference.OfType<FieldItem>())
        {
            var fieldItemValue = fieldItem.Val?.Value ?? throw PartStructureException.MissingAttribute();
            xlReference.AddFieldItem(fieldItemValue);
        }

        return xlReference;
    }

    private static void LoadPivotTableStyle(PivotTableStyle? pivotTableStyle, XLPivotTable xlPivotTable)
    {
        if (pivotTableStyle is not null)
        {
            xlPivotTable.Theme = pivotTableStyle.Name is not null && Enum.TryParse<XLPivotTableTheme>(pivotTableStyle.Name, out var xlPivotTableTheme)
                ? xlPivotTableTheme
                : XLPivotTableTheme.None;
            xlPivotTable.ShowRowHeaders = pivotTableStyle.ShowRowHeaders?.Value ?? false;
            xlPivotTable.ShowColumnHeaders = pivotTableStyle.ShowColumnHeaders?.Value ?? false;
            xlPivotTable.ShowRowStripes = pivotTableStyle.ShowRowStripes?.Value ?? false;
            xlPivotTable.ShowColumnStripes = pivotTableStyle.ShowColumnStripes?.Value ?? false;
            xlPivotTable.ShowLastColumn = pivotTableStyle.ShowColumnStripes?.Value ?? false;
        }
    }

    private static void LoadExtensionList(PivotTableDefinition pivotTable, XLPivotTable xlPivotTable)
    {
        var extList = pivotTable.GetFirstChild<PivotTableDefinitionExtensionList>();
        var ext2010 = extList?.GetFirstChild<PivotTableDefinitionExtension>();
        var ptExt2010 = ext2010?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition>();
        if (ptExt2010 is not null)
        {
            xlPivotTable.EnableCellEditing = ptExt2010.EnableEdit?.Value ?? false;
            var hideValuesRow = ptExt2010.HideValuesRow?.Value ?? false;
            xlPivotTable.ShowValuesRow = !hideValuesRow;
        }
    }
}
