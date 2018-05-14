using System;

namespace ClosedXML.Excel
{
    public enum XLNamedRangeScope
    {
        Worksheet,
        Workbook
    }

    public interface IXLNamedRange
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the comment for this named range.
        /// </summary>
        /// <value>
        /// The comment for this named range.
        /// </value>
        String Comment { get; set; }

        /// <summary>
        /// Checks if the named range contains invalid references (#REF!).
        /// </summary>
        bool IsValid { get; }

        /// <summary>
        /// Gets or sets the name of the range.
        /// </summary>
        /// <value>
        /// The name of the range.
        /// </value>
        String Name { get; set; }

        /// <summary>
        /// Gets the ranges associated with this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        IXLRanges Ranges { get; }
        String RefersTo { get; set; }

        /// <summary>
        /// Gets the scope of this named range.
        /// </summary>
        XLNamedRangeScope Scope { get; }

        /// <summary>
        /// Gets or sets the visibility of this named range.
        /// </summary>
        /// <value>
        ///   <c>true</c> if visible; otherwise, <c>false</c>.
        /// </value>
        Boolean Visible { get; set; }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Adds the specified range to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="workbook">Workbook containing the range</param>
        /// <param name="rangeAddress">The range address to add.</param>
        IXLRanges Add(XLWorkbook workbook, String rangeAddress);

        /// <summary>
        /// Adds the specified range to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="range">The range to add.</param>
        IXLRanges Add(IXLRange range);

        /// <summary>
        /// Adds the specified ranges to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="ranges">The ranges to add.</param>
        IXLRanges Add(IXLRanges ranges);

        /// <summary>
        /// Clears the list of ranges associated with this named range.
        /// <para>(it does not clear the cells)</para>
        /// </summary>
        void Clear();

        IXLNamedRange CopyTo(IXLWorksheet targetSheet);

        /// <summary>
        /// Deletes this named range (not the cells).
        /// </summary>
        void Delete();
        /// <summary>
        /// Removes the specified range from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The range address to remove.</param>
        void Remove(String rangeAddress);

        /// <summary>
        /// Removes the specified range from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="range">The range to remove.</param>
        void Remove(IXLRange range);

        /// <summary>
        /// Removes the specified ranges from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="ranges">The ranges to remove.</param>
        void Remove(IXLRanges ranges);

        IXLNamedRange SetRefersTo(String range);

        IXLNamedRange SetRefersTo(IXLRangeBase range);

        IXLNamedRange SetRefersTo(IXLRanges ranges);

        #endregion Public Methods
    }
}
