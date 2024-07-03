// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// An API for grand totals from <see cref="XLPivotTableStyleFormats"/>.
/// </summary>
internal class XLPivotStyleFormats : IXLPivotStyleFormats
{
    private readonly XLPivotTable _pivotTable;
    private readonly bool _isRowGrand;

    internal XLPivotStyleFormats(XLPivotTable pivotTable, bool isRowGrand)
    {
        _pivotTable = pivotTable;
        _isRowGrand = isRowGrand;
    }

    #region IXLPivotStyleFormats members

    public IXLPivotStyleFormat ForElement(XLPivotStyleFormatElement element)
    {
        if (element == XLPivotStyleFormatElement.None)
            throw new ArgumentException("Choose an enum value that represents an element", nameof(element));

        return GetPivotStyleFormatFor(element);
    }

    public IEnumerator<IXLPivotStyleFormat> GetEnumerator()
    {
        var elements = new[]
        {
            XLPivotStyleFormatElement.Label,
            XLPivotStyleFormatElement.Data,
            XLPivotStyleFormatElement.All,
        };

        foreach (var element in elements)
        {
            foreach (var format in _pivotTable.Formats)
            {
                if (AreaBelongsToGrandTotal(format.PivotArea, element))
                {
                    // Each pivot style format modifies all formats, so return only once per element.
                    yield return GetPivotStyleFormatFor(element);
                    break;
                }
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    #endregion IXLPivotStyleFormats members

    private XLPivotStyleFormat GetPivotStyleFormatFor(XLPivotStyleFormatElement element)
    {
        return new XLPivotStyleFormat(_pivotTable, FilterElement, ElementFactory)
        {
            AppliesTo = element
        };

        bool FilterElement(XLPivotArea pivotArea) => AreaBelongsToGrandTotal(pivotArea, element);
        XLPivotArea ElementFactory() => CreateGrandArea(element);
    }

    private bool AreaBelongsToGrandTotal(XLPivotArea area, XLPivotStyleFormatElement element)
    {
        return
            area.References.Count == 0 &&
            area.Field is null &&
            area.Type == XLPivotAreaType.Normal &&
            area.DataOnly == (element == XLPivotStyleFormatElement.Data) &&
            area.LabelOnly == (element == XLPivotStyleFormatElement.Label) &&
            area.GrandRow == _isRowGrand &&
            area.GrandCol == !_isRowGrand &&
            area.CacheIndex == false &&
            area.Offset is null &&
            !area.CollapsedLevelsAreSubtotals &&
            area.Axis is null &&
            area.FieldPosition is null;
    }

    private XLPivotArea CreateGrandArea(XLPivotStyleFormatElement element)
    {
        return new XLPivotArea
        {
            DataOnly = (element == XLPivotStyleFormatElement.Data),
            LabelOnly = (element == XLPivotStyleFormatElement.Label),
            GrandRow = _isRowGrand,
            GrandCol = !_isRowGrand,
        };
    }
}
