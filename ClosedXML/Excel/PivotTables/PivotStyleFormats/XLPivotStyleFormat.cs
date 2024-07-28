// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

internal class XLPivotStyleFormat : IXLPivotStyleFormat, IXLStylized
{
    private readonly XLPivotTable _pivotTable;
    private readonly Func<XLPivotArea, bool> _filter;
    private readonly Func<XLPivotArea> _factory;
    private XLStyleValue _styleValue;

    public XLPivotStyleFormat(IXLPivotField? field)
    {
        throw new NotImplementedException();
    }

    internal XLPivotStyleFormat(XLPivotTable pivotTable, Func<XLPivotArea, bool> filter, Func<XLPivotArea> factory)
    {
        _pivotTable = pivotTable;
        _filter = filter;
        _factory = factory;

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

    internal IList<AbstractPivotFieldReference> FieldReferences { get; } = new List<AbstractPivotFieldReference>();

    private IEnumerable<XLPivotFormat> GetFormats()
    {
        var exists = false;
        foreach (var format in _pivotTable.Formats)
        {
            if (format.Action == XLPivotFormatAction.Formatting && _filter(format.PivotArea))
            {
                exists = true;
                yield return format;
            }
        }

        if (!exists)
        {
            var format = new XLPivotFormat(_factory())
            {
                DxfStyleValue = _styleValue
            };
            _pivotTable.AddFormat(format);
            yield return format;
        }
    }
}
