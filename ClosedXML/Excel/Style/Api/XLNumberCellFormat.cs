using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

internal sealed partial class XLNumberCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLNumberCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    private int NumberFormatId
    {
        get
        {
            var numberFormat = _parent.Resolve(static x => x.NumberFormat);
            return XLPredefinedFormat.NumberFormatIds.GetValueOrDefault(numberFormat, -1);
        }
        set
        {
            if (!XLPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                throw new ArgumentOutOfRangeException($"Only predefined format is permitted. Use nested enums/members of {nameof(XLPredefinedFormat)}.");

            Format = format;
        }
    }

    private string Format
    {
        get => _parent.Resolve(static x => x.NumberFormat);
        set => _parent.ModifyNumberFormat(value);
    }

    public override bool Equals(object? obj)
    {
        return obj is IXLNumberFormatBase other && (this as IEquatable<IXLNumberFormatBase>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetNumberFormat(string numberFormat)
    {
        Format = numberFormat;
    }
}
