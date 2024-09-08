// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// A base class for pivot styling API. It has takes a selected <see cref="XLPivotArea"/>
/// and applies the style using <c>.Style*</c> API. The derived classes are responsible for
/// exposing API so user can define an area and then create the desired area (from what user
/// specified) through <see cref="GetCurrentArea"/> method.
/// </summary>
internal abstract class XLPivotStyleFormatBase : IXLPivotStyleFormat, IXLStylized
{
    protected readonly XLPivotTable PivotTable;
    private XLStyleValue _styleValue;

    protected XLPivotStyleFormatBase(XLPivotTable pivotTable)
    {
        PivotTable = pivotTable;

        // Value is Default, because it's a differential style that can't be represented yet.
        _styleValue = XLStyle.Default.Value;
    }

    #region IXLPivotStyleFormat members

    public XLPivotStyleFormatElement AppliesTo { get; init; } = XLPivotStyleFormatElement.Data;

    public IXLStyle Style
    {
        get => InnerStyle;
        set => InnerStyle = value;
    }

    #endregion IXLPivotStyleFormat members

    #region IXLStylized

    public IXLStyle InnerStyle
    {
        get => new XLStyle(this, StyleValue);
        set
        {
            var styleKey = XLStyle.GenerateKey(value);
            StyleValue = XLStyleValue.FromKey(ref styleKey);
        }
    }
    public IXLRanges RangesUsed { get; } = new XLRanges();

    public XLStyleValue StyleValue
    {
        get => _styleValue;
        set
        {
            // This sets the style of everything to the passed style, while ModifyStyle
            // is for fluent API that can modify format styles individually. Because initial
            // value of _styleValue is Default, this setter shouldn't be used as a basis
            // for modifying the DxStyleValue.
            _styleValue = value;
            foreach (var format in GetFormats())
                format.DxfStyleValue = value;
        }
    }

    public void ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
    {
        var styleKey = modification(_styleValue.Key);
        _styleValue = XLStyleValue.FromKey(ref styleKey);

        // Do not use StyleValue setter, because some formats might have different formats and
        // we should only modify them, not replace other potentially different style props of formats.
        foreach (var format in GetFormats())
        {
            var formatStyleValue = modification(format.DxfStyleValue.Key);
            format.DxfStyleValue = XLStyleValue.FromKey(ref formatStyleValue);
        }
    }

    #endregion IXLStylized

    internal abstract XLPivotArea GetCurrentArea();

    internal abstract bool Filter(XLPivotArea area);

    private IEnumerable<XLPivotFormat> GetFormats()
    {
        var exists = false;
        foreach (var format in PivotTable.Formats)
        {
            if (format.Action == XLPivotFormatAction.Formatting && Filter(format.PivotArea))
            {
                exists = true;
                yield return format;
            }
        }

        if (!exists)
        {
            var format = new XLPivotFormat(GetCurrentArea())
            {
                DxfStyleValue = _styleValue
            };
            PivotTable.AddFormat(format);
            yield return format;
        }
    }
}
