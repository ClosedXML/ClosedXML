namespace ClosedXML.Excel.IO;

/// <summary>
/// A value that should be returned when nothing should be returned, but <c>void</c> is not an option.
/// </summary>
/// <remarks>It's an empty readonly struct, so hopefully compiler can optimize it away.</remarks>
internal readonly struct Unit
{
    /// <summary>
    /// Represents the sole instance of the <see cref="Unit" /> class.
    /// </summary>
    public static Unit Value => new();
}
