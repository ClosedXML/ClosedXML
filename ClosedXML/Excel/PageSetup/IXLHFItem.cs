namespace ClosedXML.Excel
{
    public enum XLHFPredefinedText
    { 
        PageNumber, NumberOfPages, Date, Time, FullPath, Path, File, SheetName
    }
    public enum XLHFOccurrence
    { 
        AllPages, OddPages, EvenPages, FirstPage
    }

    public interface IXLHFItem: IXLWithRichString
    {
        /// <summary>
        /// Gets the text of the specified header/footer occurrence.
        /// </summary>
        /// <param name="occurrence">The occurrence.</param>
        string GetText(XLHFOccurrence occurrence);

        /// <summary>
        /// Adds the given predefined text to this header/footer item.
        /// </summary>
        /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
        IXLRichString AddText(XLHFPredefinedText predefinedText);

        /// <summary>
        /// Adds the given text to this header/footer item.
        /// </summary>
        /// <param name="text">The text to add to this header/footer item.</param>
        /// <param name="occurrence">The occurrence for the text.</param>
        IXLRichString AddText(string text, XLHFOccurrence occurrence);

        /// <summary>
        /// Adds the given predefined text to this header/footer item.
        /// </summary>
        /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
        /// <param name="occurrence">The occurrence for the predefined text.</param>
        IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence);

        /// <summary>Clears the text/formats of this header/footer item.</summary>
        /// <param name="occurrence">The occurrence to clear.</param>
        void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);

        IXLRichString AddImage(string imagePath, XLHFOccurrence occurrence = XLHFOccurrence.AllPages);
    }
}
