namespace ClosedXML.Excel;

/// <summary>
/// Absolute units of physical length.
/// </summary>
/// <remarks>
/// Pixels are relative units to the size of screen.
/// </remarks>
internal enum AbsLengthUnit
{
    Inch,
    Centimeter,
    Millimeter,

    /// <summary>
    /// 1 pt = 1/72 inch
    /// </summary>
    Point,

    /// <summary>
    /// 1 pc = 12pt.
    /// </summary>
    Pica,

    /// <summary>
    /// English metric unit.
    /// </summary>
    Emu
}
