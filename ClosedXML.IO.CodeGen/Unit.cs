namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A type to use in generics instead of <c>void</c>.
/// </summary>
internal struct Unit
{
    public static readonly Unit Value = new();
};
