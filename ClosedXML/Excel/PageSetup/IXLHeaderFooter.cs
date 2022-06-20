namespace ClosedXML.Excel
{
    public enum XLHFMode { OddPagesOnly, OddAndEvenPages, Odd }
    public interface IXLHeaderFooter
    {
        /// <summary>
        /// Gets the left header/footer item.
        /// </summary>
        IXLHFItem Left { get; }

        /// <summary>
        /// Gets the middle header/footer item.
        /// </summary>
        IXLHFItem Center { get; }

        /// <summary>
        /// Gets the right header/footer item.
        /// </summary>
        IXLHFItem Right { get; }

        /// <summary>
        /// Gets the text of the specified header/footer occurrence.
        /// </summary>
        /// <param name="occurrence">The occurrence.</param>
        string GetText(XLHFOccurrence occurrence);

        IXLHeaderFooter Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);
    }
}
