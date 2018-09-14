using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRows : IEnumerable<IXLRow>
    {
        /// <summary>
        /// Sets the height of all rows.
        /// </summary>
        /// <value>
        /// The height of all rows.
        /// </value>
        Double Height { set; }

        /// <summary>
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjusts the height of all rows based on its contents.
        /// </summary>
        IXLRows AdjustToContents();

        /// <summary>
        /// Adjusts the height of all rows based on its contents, starting from the startColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        IXLRows AdjustToContents(Int32 startColumn);

        /// <summary>
        /// Adjusts the height of all rows based on its contents, starting from the startColumn and ending at endColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        /// <param name="endColumn">The column to end calculating the row height.</param>
        IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn);

        IXLRows AdjustToContents(Double minHeight, Double maxHeight);

        IXLRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight);

        IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight);

        /// <summary>
        /// Hides all rows.
        /// </summary>
        void Hide();

        /// <summary>Unhides all rows.</summary>
        void Unhide();

        /// <summary>
        /// Increments the outline level of all rows by 1.
        /// </summary>
        void Group();

        /// <summary>
        /// Increments the outline level of all rows by 1.
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the rows will be shown collapsed.</param>
        void Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for all rows.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        void Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for all rows.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the rows will be shown collapsed.</param>
        void Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Decrements the outline level of all rows by 1.
        /// </summary>
        void Ungroup();

        /// <summary>
        /// Decrements the outline level of all rows by 1.
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove the rows from all outline levels.</param>
        void Ungroup(Boolean fromAll);

        /// <summary>
        /// Show all rows as collapsed.
        /// </summary>
        void Collapse();

        /// <summary>Expands all rows (if they're collapsed).</summary>
        void Expand();

        /// <summary>
        /// Returns the collection of cells.
        /// </summary>
        IXLCells Cells();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        /// <param name="includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCells CellsUsed(Boolean includeFormats);

        IXLCells CellsUsed(XLCellsUsedOptions options);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Adds a horizontal page break after these rows.
        /// </summary>
        IXLRows AddHorizontalPageBreaks();

        IXLRows SetDataType(XLDataType dataType);

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
