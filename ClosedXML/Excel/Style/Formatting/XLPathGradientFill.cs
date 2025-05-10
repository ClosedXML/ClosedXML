using System.Collections.Generic;

namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A path gradient fill format for a cell. The gradient has a shape of two nested rectangles,
/// where the outer rectangle is represented by cell borders and inner rectangle is defined by
/// the gradient properties (<see cref="InnerLeft"/>, <see cref="InnerRight"/>, <see cref="InnerTop"/>
/// and <see cref="InnerBottom"/>). The area inside the inner rectangle is filled by a color
/// at stop position 0 and the area between inner rectangle and outer rectangle has a gradient
/// color determined by the colors <see cref="Stops"/> collection (from color at position 0 at
/// the inner rectangle to the color at position 1 at the outer rectangle).
/// </summary>
internal record XLPathGradientFill
{
    /// <summary>
    /// The key is a position of color gradient from position 0 to 1. If stops are not available
    /// for values for 0 and 1, the gradient uses black color for them. The collection can be empty.
    /// </summary>
    public required IReadOnlyDictionary<FractionOfOne, XLColor> Stops { get; init; }

    /// <summary>
    /// A fractional position of the left side of inner rectangle inside the outer rectangle.
    /// </summary>
    public required FractionOfOne InnerLeft { get; init; }

    /// <summary>
    /// A fractional position of the right side of inner rectangle inside the outer rectangle.
    /// </summary>
    public required FractionOfOne InnerRight { get; init; }

    /// <summary>
    /// A fractional position of the top side of inner rectangle inside the outer rectangle.
    /// </summary>
    public required FractionOfOne InnerTop { get; init; }

    /// <summary>
    /// A fractional position of the bottom side of inner rectangle inside the outer rectangle.
    /// </summary>
    public required FractionOfOne InnerBottom { get; init; }
}
