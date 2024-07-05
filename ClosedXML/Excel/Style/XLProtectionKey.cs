namespace ClosedXML.Excel;

internal readonly record struct XLProtectionKey
{
    public required bool Locked { get; init; }

    public required bool Hidden { get; init; }

    public override string ToString()
    {
        return (Locked ? "Locked" : "") + (Hidden ? "Hidden" : "");
    }
}
