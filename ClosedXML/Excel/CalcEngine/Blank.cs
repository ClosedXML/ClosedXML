#nullable disable

namespace ClosedXML.Excel
{
    /// <summary>
    /// A blank value. Used as a value of blank cells or as an optional argument for function calls.
    /// </summary>
    public sealed class Blank
    {
        private Blank()
        {
        }

        /// <summary>
        /// Represents the sole instance of the <see cref="Blank" /> class.
        /// </summary>
        public static readonly Blank Value = new();

        public override string ToString() => string.Empty;
    }
}
