using System;

namespace ClosedXML.Excel;

internal partial class XLProtectionCellFormat : IXLProtection
{
    bool IXLProtection.Locked
    {
        get => Resolve(static style => style.Protection.Locked);
        set => Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    bool IXLProtection.Hidden
    {
        get => Resolve(static style => style.Protection.Hidden);
        set => Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXLStyle IXLProtection.SetLocked()
    {
        return (this as IXLProtection).SetLocked(true);
    }

    IXLStyle IXLProtection.SetLocked(bool value)
    {
        (this as IXLProtection).Locked = value;
        return _parent;
    }

    IXLStyle IXLProtection.SetHidden()
    {
        return (this as IXLProtection).SetHidden(true);
    }

    IXLStyle IXLProtection.SetHidden(bool value)
    {
        (this as IXLProtection).Hidden = value;
        return _parent;
    }

    bool IEquatable<IXLProtection>.Equals(IXLProtection? other)
    {
        if (other is null)
            return false;

        var protection = Resolve(static style => style.Protection);
        return protection.Locked == other.Locked && protection.Hidden == other.Hidden;
    }
}
