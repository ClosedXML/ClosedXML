#nullable disable

namespace ClosedXML.Graphics
{
    /// <summary>
    /// A DPI resolution.
    /// </summary>
    public readonly struct Dpi
    {
        /// <summary>
        /// Horizontal DPI resolution.
        /// </summary>
        public double X { get; }

        /// <summary>
        /// Vertical DPI resolution.
        /// </summary>
        public double Y { get; }

        public Dpi(double dpiX, double dpiY)
        {
            X = dpiX;
            Y = dpiY;
        }
    }
}
