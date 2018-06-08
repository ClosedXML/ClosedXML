using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLCells : IEnumerable<IXLCell>
    {
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

        /// <summary>
        /// Sets the type of the cells' data.
        /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
        /// </summary>
        /// <value>
        /// The type of the cell's data.
        /// </value>
        /// <exception cref="ArgumentException"></exception>
        XLDataType DataType { set; }

        IXLCells SetDataType(XLDataType dataType);

        /// <summary>
        /// Clears the contents of these cells.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLCells Clear(XLClearOptions clearOptions = XLClearOptions.All);

        /// <summary>
        /// Delete the comments of these cells.
        /// </summary>
        void DeleteComments();

        /// <summary>
        /// Delete the sparklines of these cells.
        /// </summary>
        void DeleteSparklines();

        /// <summary>
        /// Sets the cells' formula with A1 references.
        /// </summary>
        /// <value>The formula with A1 references.</value>
        String FormulaA1 { set; }

        /// <summary>
        /// Sets the cells' formula with R1C1 references.
        /// </summary>
        /// <value>The formula with R1C1 references.</value>
        String FormulaR1C1 { set; }

        IXLStyle Style { get; set; }

        void Select();
    }
}
