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
