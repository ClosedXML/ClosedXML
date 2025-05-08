namespace ClosedXML.Excel;

/// <summary>
/// A type gradient fill.
/// </summary>
internal enum XLGradientType
{
    /// <summary>
    /// Gradient fill is linear. The color transition is along a line (diagonal, horizontal...).
    /// </summary>
    Linear,

    /// <summary>
    /// Gradient fill is path. The color transition is along a rectangle, where color happens from center of rectangle outwards.
    /// </summary>
    Path
}
