namespace ClosedXML.Excel.Formatting;

internal class XLProtectionFormat
{
    public required bool Locked { get; init; }

    public required bool Hidden { get; init; }

    internal XLProtectionKey ApplyTo(XLProtectionKey protectionKey)
    {
        return new XLProtectionKey
        {
            Locked = Locked,
            Hidden = Hidden
        };
    }
}
