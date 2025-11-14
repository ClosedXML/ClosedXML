using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal sealed partial class XLProtectionCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLProtectionCellFormat(XLCellFormat parent)
    {
        _parent = parent;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXLProtection other && (this as IEquatable<IXLProtection>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXLProtection value)
    {
        Modify(static (_, other) => new XLProtectionFormatValue
        {
            Hidden = other.Hidden,
            Locked = other.Locked
        }, value);
    }

    private T Resolve<T>(Func<XLCellFormatValue, T> selector)
    {
        return _parent.Resolve(selector);
    }

    private void Modify<TProperty>(Func<XLProtectionFormatValue, TProperty, XLProtectionFormatValue> modifyProtection, TProperty value)
    {
        _parent.ModifyProtection(modifyProtection, value);
    }
}
