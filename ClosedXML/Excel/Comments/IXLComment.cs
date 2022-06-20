namespace ClosedXML.Excel
{
    public interface IXLComment : IXLFormattedText<IXLComment>, IXLDrawing<IXLComment>
    {
        /// <summary>
        /// Gets or sets this comment's author's name
        /// </summary>
        string Author { get; set; }
        /// <summary>
        /// Sets the name of the comment's author
        /// </summary>
        /// <param name="value">Author's name</param>
        IXLComment SetAuthor(string value);

        /// <summary>
        /// Adds a bolded line with the author's name
        /// </summary>
        IXLRichString AddSignature();

        void Delete();
    }

}
