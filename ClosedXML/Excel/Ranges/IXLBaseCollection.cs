using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLBaseCollection<TSingle, TMultiple> : IEnumerable<TSingle>
    {
        int Count { get; }

        IXLStyle Style { get; set; }

        IXLDataValidation SetDataValidation();

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        TMultiple AddToNamed(string rangeName);

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// </summary>
        TMultiple AddToNamed(string rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of these ranges.
        /// <para>If the named range exists, it will add these ranges to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        /// </summary>
        TMultiple AddToNamed(string rangeName, XLScope scope, string comment);

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
        object Value { set; }

        TMultiple SetValue<T>(T value);

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
        IXLCells CellsUsed(bool includeFormats);

        TMultiple SetDataType(XLDataType dataType);

        /// <summary>
        /// Clears the contents of these ranges.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        TMultiple Clear(XLClearOptions clearOptions = XLClearOptions.All);
    }
}
