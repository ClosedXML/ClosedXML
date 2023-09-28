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
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class PivotTablePartWriter
    {
        // Generates content of pivotTablePart
        internal static void GeneratePivotTablePartContent(
            WorkbookPart workbookPart,
            PivotTablePart pivotTablePart,
            XLPivotTable pt,
            XLWorkbook.SaveContext context)
        {
            var pivotSource = pt.PivotCache.CastTo<XLPivotCache>();

            var pivotTableCacheDefinitionPart = pivotTablePart.PivotTableCacheDefinitionPart;
            if (!workbookPart.GetPartById(pivotSource.WorkbookCacheRelId).Equals(pivotTableCacheDefinitionPart))
            {
                pivotTablePart.DeletePart(pivotTableCacheDefinitionPart);
                pivotTablePart.CreateRelationshipToPart(workbookPart.GetPartById(pivotSource.WorkbookCacheRelId), context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
            }

            var pti = context.PivotSources[pivotSource.Guid];

            var pivotTableDefinition = new PivotTableDefinition
            {
                Name = pt.Name,
                CacheId = pivotSource.CacheId,
                DataCaption = "Values",
                MergeItem = OpenXmlHelper.GetBooleanValue(pt.MergeAndCenterWithLabels, false),
                Indent = Convert.ToUInt32(pt.RowLabelIndent),
                PageOverThenDown = (pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown),
                PageWrap = Convert.ToUInt32(pt.FilterFieldsPageWrap),
                ShowError = String.IsNullOrEmpty(pt.ErrorValueReplacement),
                UseAutoFormatting = OpenXmlHelper.GetBooleanValue(pt.AutofitColumns, false),
                PreserveFormatting = OpenXmlHelper.GetBooleanValue(pt.PreserveCellFormatting, true),
                RowGrandTotals = OpenXmlHelper.GetBooleanValue(pt.ShowGrandTotalsRows, true),
                ColumnGrandTotals = OpenXmlHelper.GetBooleanValue(pt.ShowGrandTotalsColumns, true),
                SubtotalHiddenItems = OpenXmlHelper.GetBooleanValue(pt.FilteredItemsInSubtotals, false),
                MultipleFieldFilters = OpenXmlHelper.GetBooleanValue(pt.AllowMultipleFilters, true),
                CustomListSort = OpenXmlHelper.GetBooleanValue(pt.UseCustomListsForSorting, true),
                ShowDrill = OpenXmlHelper.GetBooleanValue(pt.ShowExpandCollapseButtons, true),
                ShowDataTips = OpenXmlHelper.GetBooleanValue(pt.ShowContextualTooltips, true),
                ShowMemberPropertyTips = OpenXmlHelper.GetBooleanValue(pt.ShowPropertiesInTooltips, true),
                ShowHeaders = OpenXmlHelper.GetBooleanValue(pt.DisplayCaptionsAndDropdowns, true),
                GridDropZones = OpenXmlHelper.GetBooleanValue(pt.ClassicPivotTableLayout, false),
                ShowEmptyRow = OpenXmlHelper.GetBooleanValue(pt.ShowEmptyItemsOnRows, false),
                ShowEmptyColumn = OpenXmlHelper.GetBooleanValue(pt.ShowEmptyItemsOnColumns, false),
                ShowItems = OpenXmlHelper.GetBooleanValue(pt.DisplayItemLabels, true),
                FieldListSortAscending = OpenXmlHelper.GetBooleanValue(pt.SortFieldsAtoZ, false),
                PrintDrill = OpenXmlHelper.GetBooleanValue(pt.PrintExpandCollapsedButtons, false),
                ItemPrintTitles = OpenXmlHelper.GetBooleanValue(pt.RepeatRowLabels, false),
                FieldPrintTitles = OpenXmlHelper.GetBooleanValue(pt.PrintTitles, false),
                EnableDrill = OpenXmlHelper.GetBooleanValue(pt.EnableShowDetails, true)
            };

            if (!String.IsNullOrEmpty(pt.ColumnHeaderCaption))
                pivotTableDefinition.ColumnHeaderCaption = StringValue.FromString(pt.ColumnHeaderCaption);

            if (!String.IsNullOrEmpty(pt.RowHeaderCaption))
                pivotTableDefinition.RowHeaderCaption = StringValue.FromString(pt.RowHeaderCaption);

            if (pt.ClassicPivotTableLayout)
            {
                pivotTableDefinition.Compact = false;
                pivotTableDefinition.CompactData = false;
            }

            if (pt.EmptyCellReplacement != null)
            {
                pivotTableDefinition.ShowMissing = true;
                pivotTableDefinition.MissingCaption = pt.EmptyCellReplacement;
            }
            else
            {
                pivotTableDefinition.ShowMissing = false;
            }

            if (pt.ErrorValueReplacement != null)
            {
                pivotTableDefinition.ShowError = true;
                pivotTableDefinition.ErrorCaption = pt.ErrorValueReplacement;
            }
            else
            {
                pivotTableDefinition.ShowError = false;
            }

            var location = new Location
            {
                FirstHeaderRow = 1U,
                FirstDataRow = 1U,
                FirstDataColumn = 1U
            };

            if (pt.ReportFilters.Any())
            {
                // Reference cell is the part BELOW the report filters
                location.Reference = pt.TargetCell.CellBelow(pt.ReportFilters.Count() + 1).Address.ToString();
            }
            else
                location.Reference = pt.TargetCell.Address.ToString();

            var rowFields = new RowFields();
            var columnFields = new ColumnFields();
            var rowItems = new RowItems();
            var columnItems = new ColumnItems();
            var pageFields = new PageFields { Count = (uint)pt.ReportFilters.Count() };
            var pivotFields = new PivotFields { Count = Convert.ToUInt32(pt.PivotCache.FieldNames.Count) };

            var orderedPageFields = new SortedDictionary<int, PageField>();
            var orderedColumnLabels = new SortedDictionary<int, Field>();
            var orderedRowLabels = new SortedDictionary<int, Field>();

            // Add value fields first
            if (pt.Values.Any())
            {
                if (pt.RowLabels.Contains(XLConstants.PivotTable.ValuesSentinalLabel))
                {
                    var f = pt.RowLabels.First(f1 => f1.SourceName == XLConstants.PivotTable.ValuesSentinalLabel);
                    orderedRowLabels.Add(pt.RowLabels.IndexOf(f), new Field { Index = -2 });
                    pivotTableDefinition.DataOnRows = true;
                }
                else if (pt.ColumnLabels.Contains(XLConstants.PivotTable.ValuesSentinalLabel))
                {
                    var f = pt.ColumnLabels.First(f1 => f1.SourceName == XLConstants.PivotTable.ValuesSentinalLabel);
                    orderedColumnLabels.Add(pt.ColumnLabels.IndexOf(f), new Field { Index = -2 });
                }
            }

            // TODO: improve performance as per https://github.com/ClosedXML/ClosedXML/pull/984#discussion_r217266491
            var fieldNames = pt.PivotCache.FieldNames;
            for (var fieldIndex = 0; fieldIndex < fieldNames.Count; ++fieldIndex)
            {
                var fieldName = fieldNames[fieldIndex];
                var ptfi = pti.Fields[fieldName];

                if (pt.RowLabels.Contains(fieldName))
                {
                    var rowLabelIndex = pt.RowLabels.IndexOf(fieldName);
                    var f = new Field { Index = fieldIndex };
                    orderedRowLabels.Add(rowLabelIndex, f);

                    if (ptfi.IsTotallyBlankField)
                        rowItems.AppendChild(new RowItem());
                    else
                    {
                        for (var i = 0; i < ptfi.DistinctValues.Count(); i++)
                        {
                            var rowItem = new RowItem();
                            rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                            rowItems.AppendChild(rowItem);
                        }
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    rowItems.AppendChild(rowItemTotal);
                }
                else if (pt.ColumnLabels.Contains(fieldName))
                {
                    var columnLabelIndex = pt.ColumnLabels.IndexOf(fieldName);
                    var f = new Field { Index = fieldIndex };
                    orderedColumnLabels.Add(columnLabelIndex, f);

                    if (ptfi.IsTotallyBlankField)
                        columnItems.AppendChild(new RowItem());
                    else
                    {
                        for (var i = 0; i < ptfi.DistinctValues.Count(); i++)
                        {
                            var rowItem = new RowItem();
                            rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                            columnItems.AppendChild(rowItem);
                        }
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    columnItems.AppendChild(rowItemTotal);
                }
            }

            for (var fieldIndex = 0; fieldIndex < fieldNames.Count; ++fieldIndex)
            {
                var fieldName = fieldNames[fieldIndex];
                var xlpf = pt.ImplementedFields.FirstOrDefault(pf => pf.SourceName.Equals(fieldName, StringComparison.OrdinalIgnoreCase));

                if (xlpf == null)
                {
                    xlpf = new XLPivotField(pt, fieldName)
                    {
                        CustomName = fieldName,
                        ShowBlankItems = true,
                    };
                }

                var ptfi = pti.Fields[fieldName];

                IXLPivotField labelOrFilterField = null;
                var pf = new PivotField
                {
                    Name = xlpf.CustomName,
                    IncludeNewItemsInFilter = OpenXmlHelper.GetBooleanValue(xlpf.IncludeNewItemsInFilter, false),
                    InsertBlankRow = OpenXmlHelper.GetBooleanValue(xlpf.InsertBlankLines, false),
                    ShowAll = OpenXmlHelper.GetBooleanValue(xlpf.ShowBlankItems, true),
                    InsertPageBreak = OpenXmlHelper.GetBooleanValue(xlpf.InsertPageBreaks, false),
                    AllDrilled = OpenXmlHelper.GetBooleanValue(xlpf.Collapsed, false),
                };
                if (!string.IsNullOrWhiteSpace(xlpf.SubtotalCaption))
                {
                    pf.SubtotalCaption = xlpf.SubtotalCaption;
                }

                if (pt.ClassicPivotTableLayout)
                {
                    pf.Outline = false;
                    pf.Compact = false;
                }
                else
                {
                    pf.Outline = OpenXmlHelper.GetBooleanValue(xlpf.Outline, true);
                    pf.Compact = OpenXmlHelper.GetBooleanValue(xlpf.Compact, true);
                }

                if (xlpf.SortType != XLPivotSortType.Default)
                {
                    pf.SortType = new EnumValue<FieldSortValues>((FieldSortValues)xlpf.SortType);
                }

                switch (pt.Subtotals)
                {
                    case XLPivotSubtotals.DoNotShow:
                        pf.DefaultSubtotal = false;
                        break;

                    case XLPivotSubtotals.AtBottom:
                        pf.SubtotalTop = false;
                        break;

                    case XLPivotSubtotals.AtTop:
                        // at top is by default
                        break;
                }

                if (xlpf.SubtotalsAtTop.HasValue)
                {
                    pf.SubtotalTop = OpenXmlHelper.GetBooleanValue(xlpf.SubtotalsAtTop.Value, true);
                }

                if (pt.RowLabels.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.RowLabels.Get(xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisRow;
                }
                else if (pt.ColumnLabels.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.ColumnLabels.Get(xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisColumn;
                }
                else if (pt.ReportFilters.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.ReportFilters.Get(xlpf.SourceName);
                    var sortOrderIndex = pt.ReportFilters.IndexOf(labelOrFilterField);

                    location.ColumnsPerPage = 1;
                    location.RowPageCount = 1;
                    pf.Axis = PivotTableAxisValues.AxisPage;

                    var pageField = new PageField
                    {
                        Hierarchy = -1,
                        Field = fieldIndex
                    };

                    if (labelOrFilterField.SelectedValues.Count == 1)
                    {
                        var selectedValue = labelOrFilterField.SelectedValues.Single();
                        var index = ptfi.DistinctValues.IndexOf(selectedValue, XLCellValueComparer.OrdinalIgnoreCase);
                        if (index >= 0)
                            pageField.Item = (UInt32)index;
                    }

                    orderedPageFields.Add(sortOrderIndex, pageField);
                }

                if ((labelOrFilterField?.SelectedValues?.Count ?? 0) > 1)
                    pf.MultipleItemSelectionAllowed = true;

                if (pt.Values.Any(p => p.SourceName == xlpf.SourceName))
                    pf.DataField = true;

                var fieldItems = new Items();

                // Output items only for row / column / filter fields
                if (!ptfi.IsTotallyBlankField &&
                    ptfi.DistinctValues.Any()
                    && (pt.RowLabels.Contains(xlpf.SourceName)
                        || pt.ColumnLabels.Contains(xlpf.SourceName)
                        || pt.ReportFilters.Contains(xlpf.SourceName)))
                {
                    uint i = 0;
                    foreach (var value in ptfi.DistinctValues)
                    {
                        var item = new Item { Index = i };

                        if (labelOrFilterField != null && labelOrFilterField.Collapsed)
                            item.HideDetails = BooleanValue.FromBoolean(false);

                        if (labelOrFilterField != null &&
                            labelOrFilterField.SelectedValues.Count > 1 &&
                            !labelOrFilterField.SelectedValues.Contains(value, XLCellValueComparer.OrdinalIgnoreCase))
                            item.Hidden = BooleanValue.FromBoolean(true);

                        fieldItems.AppendChild(item);

                        i++;
                    }
                }

                if (xlpf.Subtotals.Any())
                {
                    foreach (var subtotal in xlpf.Subtotals)
                    {
                        var itemSubtotal = new Item();
                        switch (subtotal)
                        {
                            case XLSubtotalFunction.Average:
                                pf.AverageSubTotal = true;
                                itemSubtotal.ItemType = ItemValues.Average;
                                break;

                            case XLSubtotalFunction.Count:
                                pf.CountASubtotal = true;
                                itemSubtotal.ItemType = ItemValues.CountA;
                                break;

                            case XLSubtotalFunction.CountNumbers:
                                pf.CountSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Count;
                                break;

                            case XLSubtotalFunction.Maximum:
                                pf.MaxSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Maximum;
                                break;

                            case XLSubtotalFunction.Minimum:
                                pf.MinSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Minimum;
                                break;

                            case XLSubtotalFunction.PopulationStandardDeviation:
                                pf.ApplyStandardDeviationPInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.StandardDeviationP;
                                break;

                            case XLSubtotalFunction.PopulationVariance:
                                pf.ApplyVariancePInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.VarianceP;
                                break;

                            case XLSubtotalFunction.Product:
                                pf.ApplyProductInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Product;
                                break;

                            case XLSubtotalFunction.StandardDeviation:
                                pf.ApplyStandardDeviationInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.StandardDeviation;
                                break;

                            case XLSubtotalFunction.Sum:
                                pf.SumSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Sum;
                                break;

                            case XLSubtotalFunction.Variance:
                                pf.ApplyVarianceInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Variance;
                                break;
                        }
                        fieldItems.AppendChild(itemSubtotal);
                    }
                }
                // If the field itself doesn't have subtotals, but the pivot table is set to show pivot tables, add the default item
                else if (pt.Subtotals != XLPivotSubtotals.DoNotShow)
                {
                    fieldItems.AppendChild(new Item { ItemType = ItemValues.Default });
                }

                if (fieldItems.Any())
                {
                    fieldItems.Count = Convert.ToUInt32(fieldItems.Count());
                    pf.AppendChild(fieldItems);
                }

                #region Excel 2010 Features

                if (xlpf.RepeatItemLabels)
                {
                    var pivotFieldExtensionList = new PivotFieldExtensionList();
                    pivotFieldExtensionList.RemoveNamespaceDeclaration("x");
                    var pivotFieldExtension = new PivotFieldExtension { Uri = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}" };
                    pivotFieldExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

                    var pivotField2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotField { FillDownLabels = true };

                    pivotFieldExtension.AppendChild(pivotField2);

                    pivotFieldExtensionList.AppendChild(pivotFieldExtension);
                    pf.AppendChild(pivotFieldExtensionList);
                }

                #endregion Excel 2010 Features

                pivotFields.AppendChild(pf);
            }

            pivotTableDefinition.AppendChild(location);
            pivotTableDefinition.AppendChild(pivotFields);

            if (pt.RowLabels.Any())
            {
                rowFields.Append(orderedRowLabels.Values);
                rowFields.Count = Convert.ToUInt32(rowFields.Count());
                pivotTableDefinition.AppendChild(rowFields);
            }
            else
            {
                rowItems.AppendChild(new RowItem());
            }

            if (rowItems.Any())
            {
                rowItems.Count = Convert.ToUInt32(rowItems.Count());
                pivotTableDefinition.AppendChild(rowItems);
            }

            if (pt.ColumnLabels.All(cl => cl.CustomName == XLConstants.PivotTable.ValuesSentinalLabel))
            {
                for (int i = 0; i < pt.Values.Count(); i++)
                {
                    var rowItem = new RowItem();
                    rowItem.Index = Convert.ToUInt32(i);
                    rowItem.AppendChild(new MemberPropertyIndex() { Val = i });
                    columnItems.AppendChild(rowItem);
                }
            }

            if (pt.ColumnLabels.Any())
            {
                columnFields.Append(orderedColumnLabels.Values);
                columnFields.Count = Convert.ToUInt32(columnFields.Count());
                pivotTableDefinition.AppendChild(columnFields);
            }

            if (columnItems.Any())
            {
                columnItems.Count = Convert.ToUInt32(columnItems.Count());
                pivotTableDefinition.AppendChild(columnItems);
            }

            if (pt.ReportFilters.Any())
            {
                pageFields.Append(orderedPageFields.Values);
                pageFields.Count = Convert.ToUInt32(pageFields.Count());
                pivotTableDefinition.AppendChild(pageFields);
            }

            var dataFields = new DataFields();
            foreach (var valueField in pt.Values)
            {
                // Pivot table has a field that has been removed from the source
                if (!pt.PivotCache.TryGetFieldIndex(valueField.SourceName, out var valueFieldIndex))
                {
                    continue;
                }

                UInt32 numberFormatId = 0;
                if (valueField.NumberFormat.NumberFormatId != -1 || context.SharedNumberFormats.ContainsKey(valueField.NumberFormat.NumberFormatId))
                    numberFormatId = (UInt32)valueField.NumberFormat.NumberFormatId;
                else if (context.SharedNumberFormats.Any(snf => snf.Value.NumberFormat.Format == valueField.NumberFormat.Format))
                    numberFormatId = (UInt32)context.SharedNumberFormats.First(snf => snf.Value.NumberFormat.Format == valueField.NumberFormat.Format).Key;


                var df = new DataField
                {
                    Name = valueField.CustomName,
                    Field = (UInt32)valueFieldIndex,
                    Subtotal = valueField.SummaryFormula.ToOpenXml(),
                    ShowDataAs = valueField.Calculation.ToOpenXml(),
                    NumberFormatId = numberFormatId
                };

                if (!String.IsNullOrEmpty(valueField.BaseFieldName)
                    && pt.PivotCache.TryGetFieldIndex(valueField.BaseFieldName, out var baseFieldIndex))
                {
                    df.BaseField = baseFieldIndex;

                    var items = pt.PivotCache.GetFieldValues(baseFieldIndex)
                        .Distinct()
                        .ToList();

                    var indexOfItem = items.IndexOf(valueField.BaseItemValue, XLCellValueComparer.OrdinalIgnoreCase);
                    if (indexOfItem >= 0)
                        df.BaseItem = Convert.ToUInt32(indexOfItem);
                }
                else
                {
                    df.BaseField = 0;
                }

                if (valueField.CalculationItem == XLPivotCalculationItem.Previous)
                    df.BaseItem = 1048828U;
                else if (valueField.CalculationItem == XLPivotCalculationItem.Next)
                    df.BaseItem = 1048829U;
                else if (df.BaseItem == null || !df.BaseItem.HasValue)
                    df.BaseItem = 0U;

                dataFields.AppendChild(df);
            }

            if (dataFields.Any())
            {
                dataFields.Count = Convert.ToUInt32(dataFields.Count());
                pivotTableDefinition.AppendChild(dataFields);
            }

            var pts = new PivotTableStyle
            {
                ShowRowHeaders = pt.ShowRowHeaders,
                ShowColumnHeaders = pt.ShowColumnHeaders,
                ShowRowStripes = pt.ShowRowStripes,
                ShowColumnStripes = pt.ShowColumnStripes
            };

            if (pt.Theme != XLPivotTableTheme.None)
                pts.Name = Enum.GetName(typeof(XLPivotTableTheme), pt.Theme);

            pivotTableDefinition.AppendChild(pts);

            // Pivot formats
            if (pivotTableDefinition.Formats == null)
                pivotTableDefinition.Formats = new Formats();
            else
                pivotTableDefinition.Formats.RemoveAllChildren();

            foreach (var styleFormat in pt.StyleFormats.RowGrandTotalFormats)
                GeneratePivotTableFormat(isRow: true, (XLPivotStyleFormat)styleFormat, pivotTableDefinition, context);

            foreach (var styleFormat in pt.StyleFormats.ColumnGrandTotalFormats)
                GeneratePivotTableFormat(isRow: false, (XLPivotStyleFormat)styleFormat, pivotTableDefinition, context);

            foreach (var pivotField in pt.ImplementedFields)
            {
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Header, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Header, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Subtotal, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Subtotal, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Label, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Label, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Data, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.DataValuesFormat, pivotTableDefinition, context);
            }

            if (pivotTableDefinition.Formats.Any())
            {
                pivotTableDefinition.Formats.Count = new UInt32Value((uint)pivotTableDefinition.Formats.Count());
            }
            else
                pivotTableDefinition.Formats = null;

            #region Excel 2010 Features

            var pivotTableDefinitionExtensionList = new PivotTableDefinitionExtensionList();

            var pivotTableDefinitionExtension = new PivotTableDefinitionExtension { Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
            pivotTableDefinitionExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            var pivotTableDefinition2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition
            {
                EnableEdit = pt.EnableCellEditing,
                HideValuesRow = !pt.ShowValuesRow
            };
            pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            pivotTableDefinitionExtension.AppendChild(pivotTableDefinition2);

            pivotTableDefinitionExtensionList.AppendChild(pivotTableDefinitionExtension);
            pivotTableDefinition.AppendChild(pivotTableDefinitionExtensionList);

            #endregion Excel 2010 Features

            pivotTablePart.PivotTableDefinition = pivotTableDefinition;
        }

        private static void GeneratePivotFieldFormat(XLPivotStyleFormatTarget target, XLPivotTable pt, XLPivotField pivotField, XLPivotStyleFormat styleFormat, PivotTableDefinition pivotTableDefinition, SaveContext context)
        {
            if (target == XLPivotStyleFormatTarget.GrandTotal)
                throw new ArgumentException($"Use {nameof(GeneratePivotTableFormat)} to populate grand total formats.");

            if (DefaultStyle.Equals(styleFormat.Style) || !context.DifferentialFormats.ContainsKey(((XLStyle)styleFormat.Style).Value))
                return;

            var format = new Format();

            format.FormatId = UInt32Value.FromUInt32(Convert.ToUInt32(context.DifferentialFormats[((XLStyle)styleFormat.Style).Value]));

            var pivotArea = GenerateDefaultPivotArea(target);

            pivotArea.LabelOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Label, false);
            pivotArea.DataOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Data, true);

            pivotArea.CollapsedLevelsAreSubtotals = OpenXmlHelper.GetBooleanValue(styleFormat.CollapsedLevelsAreSubtotals, false);

            if (target == XLPivotStyleFormatTarget.Header)
            {
                pivotArea.Field = pivotField.Offset;

                if (pivotField.IsOnRowAxis)
                    pivotArea.Axis = PivotTableAxisValues.AxisRow;
                else if (pivotField.IsOnColumnAxis)
                    pivotArea.Axis = PivotTableAxisValues.AxisColumn;
                else if (pivotField.IsInFilterList)
                    pivotArea.Axis = PivotTableAxisValues.AxisPage;
                else
                    throw new NotImplementedException();
            }

            //Ensure referenced pivot field is added to field references
            if (new[]
                {
                    XLPivotStyleFormatTarget.Data, XLPivotStyleFormatTarget.Label, XLPivotStyleFormatTarget.Subtotal
                }.Contains(target)
                && !styleFormat.FieldReferences.OfType<PivotLabelFieldReference>().Select(fr => fr.PivotField).Contains(pivotField))
            {
                var fr = new PivotLabelFieldReference(pivotField);
                fr.DefaultSubtotal = target == XLPivotStyleFormatTarget.Subtotal;
                styleFormat.FieldReferences.Insert(0, fr);
            }

            if (pivotArea.PivotAreaReferences == null)
                pivotArea.PivotAreaReferences = new PivotAreaReferences();
            else
                pivotArea.PivotAreaReferences.RemoveAllChildren();

            foreach (var fr in styleFormat.FieldReferences)
            {
                GeneratePivotAreaReference(pt, pivotArea.PivotAreaReferences, fr, context);
            }

            if (pivotArea.PivotAreaReferences.Any())
            {
                pivotArea.PivotAreaReferences.Count = new UInt32Value((uint)pivotArea.PivotAreaReferences.Count());
            }
            else
                pivotArea.PivotAreaReferences = null;

            format.PivotArea = pivotArea;
            pivotTableDefinition.Formats.AppendChild(format);
        }
        
        private static void GeneratePivotTableFormat(Boolean isRow, XLPivotStyleFormat styleFormat, PivotTableDefinition pivotTableDefinition, SaveContext context)
        {
            if (DefaultStyle.Equals(styleFormat.Style) || !context.DifferentialFormats.ContainsKey(((XLStyle)styleFormat.Style).Value))
                return;

            var format = new Format();

            format.FormatId = UInt32Value.FromUInt32(Convert.ToUInt32(context.DifferentialFormats[((XLStyle)styleFormat.Style).Value]));

            var pivotArea = GenerateDefaultPivotArea(XLPivotStyleFormatTarget.GrandTotal);

            pivotArea.LabelOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Label, false);
            pivotArea.DataOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Data, true);

            pivotArea.GrandColumn = OpenXmlHelper.GetBooleanValue(!isRow, false);
            pivotArea.GrandRow = OpenXmlHelper.GetBooleanValue(isRow, false);
            pivotArea.Axis = isRow ? PivotTableAxisValues.AxisRow : PivotTableAxisValues.AxisColumn;

            format.PivotArea = pivotArea;

            pivotTableDefinition.Formats.AppendChild(format);
        }

        private static PivotArea GenerateDefaultPivotArea(XLPivotStyleFormatTarget target)
        {
            switch (target)
            {
                case XLPivotStyleFormatTarget.Header:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Button,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(true, false),
                        Outline = OpenXmlHelper.GetBooleanValue(false, true),
                    };

                case XLPivotStyleFormatTarget.Subtotal:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                    };

                case XLPivotStyleFormatTarget.GrandTotal:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(false, false),
                    };

                case XLPivotStyleFormatTarget.Label:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(true, false),
                    };

                case XLPivotStyleFormatTarget.Data:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                    };

                default:
                    throw new NotImplementedException();
            }
        }

        private static void GeneratePivotAreaReference(XLPivotTable pt, PivotAreaReferences pivotAreaReferences, AbstractPivotFieldReference fieldReference, SaveContext context)
        {
            var pivotAreaReference = new PivotAreaReference();

            pivotAreaReference.DefaultSubtotal = OpenXmlHelper.GetBooleanValue(fieldReference.DefaultSubtotal, false);
            pivotAreaReference.Field = fieldReference.GetFieldOffset();

            var pivotSource = pt.PivotCache.CastTo<XLPivotCache>();
            var matchedOffsets = fieldReference.Match(context.PivotSources[pivotSource.Guid], pt);
            foreach (var o in matchedOffsets)
            {
                pivotAreaReference.AppendChild(new FieldItem { Val = UInt32Value.FromUInt32((uint)o) });
            }

            pivotAreaReferences.AppendChild(pivotAreaReference);
        }
    }
}
