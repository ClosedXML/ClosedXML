using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;
using ClosedXML.Extensions;
using DocumentFormat.OpenXml.Packaging;
using static ClosedXML.Excel.IO.OpenXmlConst;
using static ClosedXML.Excel.XLWorkbook;
using Array = System.Array;

namespace ClosedXML.Excel.IO;

internal class PivotTableDefinitionPartWriter2
{
    internal static void WriteContent(PivotTablePart pivotTablePart, XLPivotTable pt, SaveContext context)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = XLHelper.NoBomUTF8
        };

        using var partStream = pivotTablePart.GetStream(FileMode.Create);
        using var xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();
        xml.WriteStartElement("pivotTableDefinition", Main2006SsNs);
        xml.WriteAttributeString("xmlns", "mc", null, MarkupCompatibilityNs);

        // Mark revision as ignorable extension
        xml.WriteAttributeString("mc", "Ignorable", null, "xr");
        xml.WriteAttributeString("xmlns", "xr", null, RevisionNs);
        xml.WriteAttribute("name", pt.Name);
        xml.WriteAttribute("cacheId", pt.PivotCache.CacheId!.Value); // TODO: Maybe not nullable?
        xml.WriteAttributeDefault("dataOnRows", pt.DataOnRows, false);
        xml.WriteAttributeOptional("dataPosition", pt.DataPosition);
        xml.WriteAttributeOptional("autoFormatId", pt.AutoFormatId);

        // Although apply*Formats do have default value `false`, Excel always writes them.
        xml.WriteAttribute("applyNumberFormats", pt.ApplyNumberFormats);
        xml.WriteAttribute("applyBorderFormats", pt.ApplyBorderFormats);
        xml.WriteAttribute("applyFontFormats", pt.ApplyFontFormats);
        xml.WriteAttribute("applyPatternFormats", pt.ApplyPatternFormats);
        xml.WriteAttribute("applyAlignmentFormats", pt.ApplyAlignmentFormats);
        xml.WriteAttribute("applyWidthHeightFormats", pt.ApplyWidthHeightFormats);

        xml.WriteAttribute("dataCaption", pt.DataCaption);
        xml.WriteAttributeOptional("grandTotalCaption", pt.GrandTotalCaption);
        xml.WriteAttributeOptional("errorCaption", pt.ErrorValueReplacement);
        xml.WriteAttributeDefault("showError", pt.ShowError, false);
        xml.WriteAttributeOptional("missingCaption", pt.MissingCaption);
        xml.WriteAttributeDefault("showMissing", pt.ShowMissing, true);
        xml.WriteAttributeOptional("pageStyle", pt.PageStyle);
        xml.WriteAttributeOptional("pivotTableStyle", pt.PivotTableStyleName);
        xml.WriteAttributeOptional("vacatedStyle", pt.VacatedStyle);
        xml.WriteAttributeOptional("tag", pt.Tag);
        xml.WriteAttributeDefault("updatedVersion", pt.UpdatedVersion, 0);
        xml.WriteAttributeDefault("minRefreshableVersion", pt.MinRefreshableVersion, 0);
        xml.WriteAttributeDefault("asteriskTotals", pt.AsteriskTotals, false);
        xml.WriteAttributeDefault("showItems", pt.DisplayItemLabels, true);
        xml.WriteAttributeDefault("editData", pt.EditData, false);
        xml.WriteAttributeDefault("disableFieldList", pt.DisableFieldList, false);
        xml.WriteAttributeDefault(@"showCalcMbrs", pt.ShowCalculatedMembers, true);
        xml.WriteAttributeDefault("visualTotals", pt.VisualTotals, true);
        xml.WriteAttributeDefault("showMultipleLabel", pt.ShowMultipleLabel, true);
        xml.WriteAttributeDefault("showDataDropDown", pt.ShowDataDropDown, true);
        xml.WriteAttributeDefault("showDrill", pt.ShowExpandCollapseButtons, true);
        xml.WriteAttributeDefault("printDrill", pt.PrintExpandCollapsedButtons, false);
        xml.WriteAttributeDefault("showMemberPropertyTips", pt.ShowPropertiesInTooltips, true);
        xml.WriteAttributeDefault("showDataTips", pt.ShowContextualTooltips, true);
        xml.WriteAttributeDefault("enableWizard", pt.EnableEditingMechanism, true);
        xml.WriteAttributeDefault("enableDrill", pt.EnableShowDetails, true);
        xml.WriteAttributeDefault("enableFieldProperties", pt.EnableFieldProperties, true);
        xml.WriteAttributeDefault("preserveFormatting", pt.PreserveCellFormatting, true);
        xml.WriteAttributeDefault("useAutoFormatting", pt.AutofitColumns, false);
        xml.WriteAttributeDefault("pageWrap", checked((uint)pt.FilterFieldsPageWrap), 0);
        xml.WriteAttributeDefault("pageOverThenDown", pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown, false);
        xml.WriteAttributeDefault("subtotalHiddenItems", pt.FilteredItemsInSubtotals, false);
        xml.WriteAttributeDefault("rowGrandTotals", pt.ShowGrandTotalsRows, true);
        xml.WriteAttributeDefault("colGrandTotals", pt.ShowGrandTotalsColumns, true);
        xml.WriteAttributeDefault("fieldPrintTitles", pt.PrintTitles, false);
        xml.WriteAttributeDefault("itemPrintTitles", pt.RepeatRowLabels, false);
        xml.WriteAttributeDefault("mergeItem", pt.MergeAndCenterWithLabels, false);
        xml.WriteAttributeDefault("showDropZones", pt.ShowDropZones, true);
        xml.WriteAttributeDefault("createdVersion", pt.PivotCacheCreatedVersion, 0);
        xml.WriteAttributeDefault("indent", checked((uint)pt.RowLabelIndent), 1);
        xml.WriteAttributeDefault("showEmptyRow", pt.ShowEmptyItemsOnRows, false);
        xml.WriteAttributeDefault("showEmptyCol", pt.ShowEmptyItemsOnColumns, false);
        xml.WriteAttributeDefault("showHeaders", pt.DisplayCaptionsAndDropdowns, true);
        xml.WriteAttributeDefault("compact", pt.Compact, true);
        xml.WriteAttributeDefault("outline", pt.Outline, false);
        xml.WriteAttributeDefault("outlineData", pt.OutlineData, false);
        xml.WriteAttributeDefault("compactData", pt.CompactData, true);
        xml.WriteAttributeDefault("published", pt.Published, false);
        xml.WriteAttributeDefault("gridDropZones", pt.ClassicPivotTableLayout, false);
        xml.WriteAttributeDefault("immersive", pt.StopImmersiveUi, true);
        xml.WriteAttributeDefault("multipleFieldFilters", pt.AllowMultipleFilters, true);
        xml.WriteAttributeDefault("chartFormat", pt.ChartFormat, 0);
        xml.WriteAttributeOptional("rowHeaderCaption", pt.RowHeaderCaption);
        xml.WriteAttributeOptional("colHeaderCaption", pt.ColumnHeaderCaption);
        xml.WriteAttributeDefault("fieldListSortAscending", pt.SortFieldsAtoZ, false);
        xml.WriteAttributeDefault(@"mdxSubqueries", pt.MdxSubQueries, false);
        xml.WriteAttributeDefault("customListSort", pt.UseCustomListsForSorting, true);

        // Location
        xml.WriteStartElement("location", Main2006SsNs);
        xml.WriteAttribute("ref", pt.Area.ToString());
        xml.WriteAttribute("firstHeaderRow", pt.FirstHeaderRow);
        xml.WriteAttribute("firstDataRow", pt.FirstDataRow);
        xml.WriteAttribute("firstDataCol", pt.FirstDataCol);
        xml.WriteAttributeDefault("rowPageCount", pt.RowPageCount, 0);
        xml.WriteAttributeDefault("colPageCount", pt.ColumnPageCount, 0);
        xml.WriteEndElement(); // location

        // Pivot Fields
        xml.WriteStartElement("pivotFields", Main2006SsNs);
        xml.WriteAttribute("count", pt.PivotFields.Count);

        foreach (var pf in pt.PivotFields)
        {
            xml.WriteStartElement("pivotField", Main2006SsNs);
            xml.WriteAttributeOptional("name", pf.Name);

            if (pf.Axis is not null)
            {
                var axisAttr = GetAxisAttr(pf.Axis.Value);
                xml.WriteAttribute("axis", axisAttr);
            }

            xml.WriteAttributeDefault("dataField", pf.DataField, false);
            xml.WriteAttributeOptional("subtotalCaption", pf.SubtotalCaption);
            xml.WriteAttributeDefault("showDropDowns", pf.ShowDropDowns, true);
            xml.WriteAttributeDefault("hiddenLevel", pf.HiddenLevel, false);
            xml.WriteAttributeOptional("uniqueMemberProperty", pf.UniqueMemberProperty);
            xml.WriteAttributeDefault("compact", pf.Compact, true);
            xml.WriteAttributeDefault("allDrilled", pf.AllDrilled, false);
            xml.WriteAttributeOptional("numFmtId", pf.NumberFormatId);
            xml.WriteAttributeDefault("outline", pf.Outline, true);
            xml.WriteAttributeDefault("subtotalTop", pf.SubtotalTop, true);
            xml.WriteAttributeDefault("dragToRow", pf.DragToRow, true);
            xml.WriteAttributeDefault("dragToCol", pf.DragToColumn, true);
            xml.WriteAttributeDefault("multipleItemSelectionAllowed", pf.MultipleItemSelectionAllowed, false);
            xml.WriteAttributeDefault("dragToPage", pf.DragToPage, true);
            xml.WriteAttributeDefault("dragToData", pf.DragToData, true);
            xml.WriteAttributeDefault("dragOff", pf.DragOff, true);
            xml.WriteAttributeDefault("showAll", pf.ShowAll, true);
            xml.WriteAttributeDefault("insertBlankRow", pf.InsertBlankRow, false);
            xml.WriteAttributeDefault("serverField", pf.ServerField, false);
            xml.WriteAttributeDefault("insertPageBreak", pf.InsertPageBreak, false);
            xml.WriteAttributeDefault("autoShow", pf.AutoShow, false);
            xml.WriteAttributeDefault("topAutoShow", pf.TopAutoShow, true);
            xml.WriteAttributeDefault("hideNewItems", pf.HideNewItems, false);
            xml.WriteAttributeDefault("measureFilter", pf.MeasureFilter, false);
            xml.WriteAttributeDefault("includeNewItemsInFilter", pf.IncludeNewItemsInFilter, false);
            xml.WriteAttributeDefault("itemPageCount", pf.ItemPageCount, 10);
            if (pf.SortType != XLPivotSortType.Default)
            {
                var sortTypeAttr = pf.SortType switch
                {
                    XLPivotSortType.Default => "manual",
                    XLPivotSortType.Ascending => "ascending",
                    XLPivotSortType.Descending => "descending",
                    _ => throw new UnreachableException(),
                };
                xml.WriteAttribute("sortType", sortTypeAttr);
            }

            xml.WriteAttributeOptional("dataSourceSort", pf.DataSourceSort);
            xml.WriteAttributeDefault("nonAutoSortDefault", pf.NonAutoSortDefault, false);
            xml.WriteAttributeOptional("rankBy", pf.RankBy);
            xml.WriteAttributeDefault("defaultSubtotal", pf.DefaultSubtotal, true);
            xml.WriteAttributeDefault("sumSubtotal", pf.SumSubtotal, false);
            xml.WriteAttributeDefault("countASubtotal", pf.CountASubtotal, false);
            xml.WriteAttributeDefault("avgSubtotal", pf.AvgSubtotal, false);
            xml.WriteAttributeDefault("maxSubtotal", pf.MaxSubtotal, false);
            xml.WriteAttributeDefault("minSubtotal", pf.MinSubtotal, false);
            xml.WriteAttributeDefault("productSubtotal", pf.ProductSubtotal, false);
            xml.WriteAttributeDefault("countSubtotal", pf.CountSubtotal, false);
            xml.WriteAttributeDefault("stdDevSubtotal", pf.StdDevSubtotal, false);
            xml.WriteAttributeDefault("stdDevPSubtotal", pf.StdDevPSubtotal, false);
            xml.WriteAttributeDefault("varSubtotal", pf.VarSubtotal, false);
            xml.WriteAttributeDefault("varPSubtotal", pf.VarPSubtotal, false);
            xml.WriteAttributeDefault("showPropCell", pf.ShowPropCell, false);
            xml.WriteAttributeDefault("showPropTip", pf.ShowPropTip, false);
            xml.WriteAttributeDefault("showPropAsCaption", pf.ShowPropAsCaption, false);
            xml.WriteAttributeDefault("defaultAttributeDrillState", pf.DefaultAttributeDrillState, false);

            // items
            if (pf.Items.Count > 0)
            {
                xml.WriteStartElement("items", Main2006SsNs);
                xml.WriteAttribute("count", pf.Items.Count);
                foreach (var pfItem in pf.Items)
                {
                    xml.WriteStartElement("item", Main2006SsNs);
                    xml.WriteAttributeOptional("n", pfItem.ItemUserCaption);
                    if (pfItem.ItemType != XLPivotItemType.Data)
                    {
                        var itemTypeAttr = GetItemTypeAttr(pfItem.ItemType);
                        xml.WriteAttribute("t", itemTypeAttr);
                    }

                    xml.WriteAttributeDefault("h", pfItem.Hidden, false);
                    xml.WriteAttributeDefault("s", pfItem.ValueIsString, false);
                    xml.WriteAttributeDefault("sd", pfItem.HideDetails, true);
                    xml.WriteAttributeDefault("f", pfItem.CalculatedMember, false);
                    xml.WriteAttributeDefault("m", pfItem.Missing, false);
                    xml.WriteAttributeDefault("c", pfItem.ApproximatelyHasChildren, false);
                    xml.WriteAttributeOptional("x", pfItem.ItemIndex);
                    xml.WriteAttributeDefault("d", pfItem.IsExpanded, false);
                    xml.WriteAttributeDefault("e", pfItem.DrillAcrossAttributes, true);
                    xml.WriteEndElement(); // item
                }

                xml.WriteEndElement(); // items
            }

            // TODO: autoSortScope, but not yet represented.

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // pivotFields

        WriteAxis(xml, pt.RowAxis, "rowFields", "rowItems");
        WriteAxis(xml, pt.ColumnAxis, "colFields", "colItems");

        if (pt.PageFields.Count > 0)
        {
            xml.WriteStartElement("pageFields", Main2006SsNs);
            xml.WriteAttribute("count", pt.PageFields.Count);
            foreach (var pageField in pt.PageFields)
            {
                xml.WriteStartElement("pageField", Main2006SsNs);
                xml.WriteAttribute("fld", pageField.Field);
                xml.WriteAttributeOptional("item", pageField.ItemIndex);
                xml.WriteAttributeOptional("hier", pageField.HierarchyIndex);
                xml.WriteAttributeOptional("name", pageField.HierarchyUniqueName);
                xml.WriteAttributeOptional("cap", pageField.HierarchyDisplayName);
                xml.WriteEndElement(); // pageField
            }

            xml.WriteEndElement(); // pageFields
        }

        if (pt.DataFields.Count > 0)
        {
            xml.WriteStartElement("dataFields", Main2006SsNs);
            xml.WriteAttribute("count", pt.DataFields.Count);
            foreach (var dataField in pt.DataFields)
            {
                xml.WriteStartElement("dataField", Main2006SsNs);
                xml.WriteAttributeOptional("name", dataField.DataFieldName);
                xml.WriteAttribute("fld", dataField.Field);
                if (dataField.Subtotal != XLPivotSummary.Sum)
                {
                    var subtotalAttr = dataField.Subtotal switch
                    {
                        XLPivotSummary.Sum => "sum",
                        XLPivotSummary.Count => "count",
                        XLPivotSummary.Average => "average",
                        XLPivotSummary.Minimum => "min",
                        XLPivotSummary.Maximum => "max",
                        XLPivotSummary.Product => "product",
                        XLPivotSummary.CountNumbers => "countNums",
                        XLPivotSummary.StandardDeviation => "stdDev",
                        XLPivotSummary.PopulationStandardDeviation => "stdDevp",
                        XLPivotSummary.Variance => "var",
                        XLPivotSummary.PopulationVariance => "varp",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("subtotal", subtotalAttr);
                }

                if (dataField.ShowDataAsFormat != XLPivotCalculation.Normal)
                {
                    var showDataAsAttr = dataField.ShowDataAsFormat switch
                    {
                        XLPivotCalculation.Normal => "normal",
                        XLPivotCalculation.DifferenceFrom => "difference",
                        XLPivotCalculation.PercentageOf => "percent",
                        XLPivotCalculation.PercentageDifferenceFrom => "percentDiff",
                        XLPivotCalculation.RunningTotal => "runTotal",
                        XLPivotCalculation.PercentageOfRow => "percentOfRow",
                        XLPivotCalculation.PercentageOfColumn => "percentOfCol",
                        XLPivotCalculation.PercentageOfTotal => "percentOfTotal",
                        XLPivotCalculation.Index => "index",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("showDataAs", showDataAsAttr);
                }

                xml.WriteAttributeDefault("baseField", dataField.BaseField, -1);
                xml.WriteAttributeDefault("baseItem", dataField.BaseItem, 1048832);
                xml.WriteAttributeOptional("numFmtId", dataField.NumberFormatId);
                xml.WriteEndElement(); // dataField
            }

            xml.WriteEndElement(); // dataFields
        }

        if (pt.Formats.Count > 0)
        {
            xml.WriteStartElement("formats", Main2006SsNs);
            xml.WriteAttribute("count", pt.Formats.Count);
            foreach (var format in pt.Formats)
            {
                xml.WriteStartElement("format", Main2006SsNs);
                if (format.Action != XLPivotFormatAction.Formatting)
                {
                    var actionAttr = format.Action switch
                    {
                        XLPivotFormatAction.Blank => "blank",
                        XLPivotFormatAction.Formatting => "formatting",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("action", actionAttr);
                }

                // DxfId is optional.
                if (format.DxfStyle.Value != XLStyleValue.Default)
                {
                    var dxfId = context.DifferentialFormats[format.DxfStyle.Value];
                    xml.WriteAttribute("dxfId", dxfId);
                }

                var pivotArea = format.PivotArea;
                WritePivotArea(xml, pivotArea);
                xml.WriteEndElement(); // format
            }
            xml.WriteEndElement(); // formats
        }

        xml.WriteEndElement(); // pivotTableDefinition
    }

    private static void WriteAxis(XmlWriter xml, XLPivotTableAxis axis, string fieldsElement, string itemsElement)
    {
        if (axis.Fields.Count > 0)
        {
            xml.WriteStartElement(fieldsElement, Main2006SsNs);
            xml.WriteAttribute("count", axis.Fields.Count);
            foreach (var axisField in axis.Fields)
            {
                xml.WriteStartElement("field", Main2006SsNs);
                xml.WriteAttribute("x", axisField.Value);
                xml.WriteEndElement();
            }

            xml.WriteEndElement(); // rowFields
        }

        if (axis.Items.Count > 0)
        {
            xml.WriteStartElement(itemsElement, Main2006SsNs);
            xml.WriteAttribute("count", axis.Items.Count);

            IReadOnlyList<int> previousItems = Array.Empty<int>();
            foreach (var axisItem in axis.Items)
            {
                xml.WriteStartElement("i", Main2006SsNs);
                if (axisItem.ItemType != XLPivotItemType.Data)
                {
                    var itemTypeAttr = GetItemTypeAttr(axisItem.ItemType);
                    xml.WriteAttribute("t", itemTypeAttr);
                }

                // 'r' attribute means repeat data from previous axis item.
                var maxLen = Math.Min(previousItems.Count, axisItem.FieldItem.Count);
                var r = 0;
                while (r < maxLen && previousItems[r] == axisItem.FieldItem[r])
                    r++;

                xml.WriteAttributeDefault("r", r, 0);
                xml.WriteAttributeDefault("i", axisItem.DataItem, 0); // Data field index

                foreach (var fieldItem in axisItem.FieldItem)
                {
                    xml.WriteStartElement("x", Main2006SsNs);
                    xml.WriteAttributeDefault("v", fieldItem, 0);
                    xml.WriteEndElement(); // x
                }

                xml.WriteEndElement(); // i
                previousItems = axisItem.FieldItem;
            }

            xml.WriteEndElement();
        }
    }

    private static void WritePivotArea(XmlWriter xml, XLPivotArea pivotArea)
    {
        xml.WriteStartElement("pivotArea", Main2006SsNs);
        xml.WriteAttributeOptional("field", pivotArea.Field?.Value);
        if (pivotArea.Type != XLPivotAreaType.Normal)
        {
            var typeAttr = pivotArea.Type switch
            {
                XLPivotAreaType.None => "none",
                XLPivotAreaType.Normal => "normal",
                XLPivotAreaType.Data => "data",
                XLPivotAreaType.All => "all",
                XLPivotAreaType.Origin => "origin",
                XLPivotAreaType.Button => "button",
                XLPivotAreaType.TopRight => "topRight",
                XLPivotAreaType.TopEnd => "topEnd",
                _ => throw new UnreachableException(),
            };
            xml.WriteAttribute("type", typeAttr);
        }

        xml.WriteAttributeDefault("dataOnly", pivotArea.DataOnly, true);
        xml.WriteAttributeDefault("labelOnly", pivotArea.LabelOnly, false);
        xml.WriteAttributeDefault("grandRow", pivotArea.GrandRow, false);
        xml.WriteAttributeDefault("grandCol", pivotArea.GrandCol, false);
        xml.WriteAttributeDefault("cacheIndex", pivotArea.CacheIndex, false);
        xml.WriteAttributeDefault("outline", pivotArea.Outline, true);
        if (pivotArea.Offset is not null)
            xml.WriteAttribute("offset", pivotArea.Offset.ToString());

        xml.WriteAttributeDefault("collapsedLevelsAreSubtotals", pivotArea.CollapsedLevelsAreSubtotals, false);
        if (pivotArea.Axis is not null)
            xml.WriteAttribute("axis", GetAxisAttr(pivotArea.Axis.Value));

        xml.WriteAttributeOptional("fieldPosition", pivotArea.FieldPosition);

        if (pivotArea.References.Count > 0)
        {
            xml.WriteStartElement("references", Main2006SsNs);
            xml.WriteAttribute("count", pivotArea.References.Count);
            foreach (var reference in pivotArea.References)
            {
                xml.WriteStartElement("reference", Main2006SsNs);
                xml.WriteAttributeOptional("field", reference.Field);
                xml.WriteAttribute("count", reference.FieldItems.Count);
                xml.WriteAttributeDefault("selected", reference.Selected, true);
                xml.WriteAttributeDefault("byPosition", reference.ByPosition, false);
                xml.WriteAttributeDefault("relative", reference.Relative, false);
                xml.WriteAttributeDefault("defaultSubtotal", reference.DefaultSubtotal, false);
                xml.WriteAttributeDefault("sumSubtotal", reference.SumSubtotal, false);
                xml.WriteAttributeDefault("countASubtotal", reference.CountASubtotal, false);
                xml.WriteAttributeDefault("avgSubtotal", reference.AvgSubtotal, false);
                xml.WriteAttributeDefault("maxSubtotal", reference.MaxSubtotal, false);
                xml.WriteAttributeDefault("minSubtotal", reference.MinSubtotal, false);
                xml.WriteAttributeDefault("productSubtotal", reference.ProductSubtotal, false);
                xml.WriteAttributeDefault("countSubtotal", reference.CountSubtotal, false);
                xml.WriteAttributeDefault("stdDevSubtotal", reference.StdDevSubtotal, false);
                xml.WriteAttributeDefault("stdDevPSubtotal", reference.StdDevPSubtotal, false);
                xml.WriteAttributeDefault("varSubtotal", reference.VarSubtotal, false);
                xml.WriteAttributeDefault("varPSubtotal", reference.VarPSubtotal, false);

                foreach (var fieldItem in reference.FieldItems)
                {
                    xml.WriteStartElement("x", Main2006SsNs);
                    xml.WriteAttribute("v", fieldItem);
                    xml.WriteEndElement(); // x
                }

                xml.WriteEndElement(); // reference
            }

            xml.WriteEndElement(); // references
        }

        xml.WriteEndElement(); // pivotArea
    }

    private static string GetItemTypeAttr(XLPivotItemType itemType)
    {
        var itemTypeAttr = itemType switch
        {
            XLPivotItemType.Avg => "avg",
            XLPivotItemType.Blank => "blank",
            XLPivotItemType.Count => "count",
            XLPivotItemType.CountA => "countA",
            XLPivotItemType.Data => "data",
            XLPivotItemType.Default => "default",
            XLPivotItemType.Grand => "grand",
            XLPivotItemType.Max => "max",
            XLPivotItemType.Min => "min",
            XLPivotItemType.Product => "product",
            XLPivotItemType.StdDev => "stdDev",
            XLPivotItemType.StdDevP => "stdDevP",
            XLPivotItemType.Sum => "sum",
            XLPivotItemType.Var => "var",
            XLPivotItemType.VarP => "varP",
            _ => throw new UnreachableException(),
        };
        return itemTypeAttr;
    }

    private static string GetAxisAttr(XLPivotAxis axis)
    {
        return axis switch
        {
            XLPivotAxis.AxisRow => "axisRow",
            XLPivotAxis.AxisCol => "axisCol",
            XLPivotAxis.AxisPage => "axisPage",
            XLPivotAxis.AxisValues => "axisValues",
            _ => throw new UnreachableException(),
        };
    }
}
