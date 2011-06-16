using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRanges: IEnumerable<IXLRange>
    {
        /// <summary>
        /// Clears the contents of the ranges (including styles).
        /// </summary>
        void Clear();
        /// <summary>
        /// Adds the specified range to this group.
        /// </summary>
        /// <param name="range">The range to add to this group.</param>
        void Add(IXLRange range);
        /// <summary>
        /// Removes the specified range from this group.
        /// </summary>
        /// <param name="range">The range to remove from this group.</param>
        void Remove(IXLRange range);

        Int32 Count { get; }

        Boolean Contains(IXLRange range);

        IXLStyle Style { get; set; }

        IXLDataValidation DataValidation { get; }

        /// <summary>
        /// Creates a named range out of these ranges. 
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLRanges AddToNamed(String rangeName);

        /// <summary>
        /// Creates a named range out of these ranges. 
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        IXLRanges AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of these ranges. 
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        IXLRanges AddToNamed(String rangeName, XLScope scope, String comment);

        /// <summary>
        /// Sets the cells' value.
        /// <para>If the object is an IEnumerable ClosedXML will copy the collection's data into a table starting from each cell.</para>
        /// <para>If the object is a range ClosedXML will copy the range starting from each cell.</para>
        /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
        /// <para>ClosedXML will try to translate it to the corresponding type, if it can't then the value will be left as a string.</para>
        /// </summary>
        /// <value>
        /// The object containing the value(s) to set.
        /// </value>
        Object Value { set; }

        IXLRanges SetValue<T>(T value);

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
        /// <param name="includeStyles">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeStyles);

        IXLRanges SetDataType(XLCellValues dataType);
    }
}
