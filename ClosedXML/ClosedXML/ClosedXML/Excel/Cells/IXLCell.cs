using System;
using System.Collections;

namespace ClosedXML.Excel
{
    public enum XLCellValues { Text, Number, Boolean, DateTime, TimeSpan }

    public interface IXLCell
    {
        /// <summary>
        /// Gets or sets the cell's value. To get a strongly typed object use the method GetValue&lt;T&gt;.
        /// <para>If the object is an IEnumerable ClosedXML will copy the collection's data into a table starting from this cell.</para>
        /// <para>If the object is a range ClosedXML will copy the range starting from this cell.</para>
        /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
        /// <para>ClosedXML will try to translate it to the corresponding type, if it can't then the value will be left as a string.</para>
        /// </summary>
        /// <value>
        /// The object containing the value(s) to set.
        /// </value>
        Object Value { get; set; }

        /// <summary>Gets this cell's address, relative to the worksheet.</summary>
        /// <value>The cell's address.</value>
        IXLAddress Address { get;  }

        /// <summary>
        /// Gets or sets the type of this cell's data.
        /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
        /// </summary>
        /// <value>
        /// The type of the cell's data.
        /// </value>
        /// <exception cref="ArgumentException"></exception>
        XLCellValues DataType { get; set; }

        IXLCell SetDataType(XLCellValues dataType);

        IXLCell SetValue<T>(T value);

        /// <summary>
        /// Gets the cell's value converted to the T type.
        /// <para>ClosedXML will try to covert the current value to the T type.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to the T type.</para>
        /// </summary>
        /// <typeparam name="T">The return type.</typeparam>
        /// <exception cref="ArgumentException"></exception>
        T GetValue<T>();

        /// <summary>
        /// Gets the cell's value converted to a String.
        /// </summary>
        String GetString();

        /// <summary>
        /// Gets the cell's value formatted depending on the cell's data type and style.
        /// </summary>
        String GetFormattedString();

        /// <summary>
        /// Gets the cell's value converted to Double.
        /// <para>ClosedXML will try to covert the current value to Double.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to Double.</para>
        /// </summary>
        Double GetDouble();

        /// <summary>
        /// Gets the cell's value converted to Boolean.
        /// <para>ClosedXML will try to covert the current value to Boolean.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to Boolean.</para>
        /// </summary>
        Boolean GetBoolean();

        /// <summary>
        /// Gets the cell's value converted to DateTime.
        /// <para>ClosedXML will try to covert the current value to DateTime.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to DateTime.</para>
        /// </summary>
        DateTime GetDateTime();

        /// <summary>
        /// Gets the cell's value converted to TimeSpan.
        /// <para>ClosedXML will try to covert the current value to TimeSpan.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to TimeSpan.</para>
        /// </summary>
        TimeSpan GetTimeSpan();

        IXLRichText GetRichText();

        /// <summary>
        /// Clears the contents of this cell (including styles).
        /// </summary>
        void Clear();

        /// <summary>
        /// Clears the styles of this cell (preserving number formats).
        /// </summary>
        void ClearStyles();

        /// <summary>
        /// Deletes the current cell and shifts the surrounding cells according to the shiftDeleteCells parameter.
        /// </summary>
        /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
        void Delete(XLShiftDeletedCells shiftDeleteCells);

        /// <summary>
        /// Gets or sets the cell's formula with A1 references.
        /// </summary>
        /// <value>The formula with A1 references.</value>
        String FormulaA1 { get; set; }

        /// <summary>
        /// Gets or sets the cell's formula with R1C1 references.
        /// </summary>
        /// <value>The formula with R1C1 references.</value>
        String FormulaR1C1 { get; set; }

        /// <summary>
        /// Returns this cell as an IXLRange.
        /// </summary>
        IXLRange AsRange();

        IXLStyle Style { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this cell's text should be shared or not.
        /// </summary>
        /// <value>
        ///   If false the cell's text will not be shared and stored as an inline value.
        /// </value>
        Boolean ShareString { get; set; }

        IXLRange InsertData(IEnumerable data);
        IXLTable InsertTable(IEnumerable data);
        IXLTable InsertTable(IEnumerable data, Boolean createTable);
        IXLTable InsertTable(IEnumerable data, String tableName);
        IXLTable InsertTable(IEnumerable data, String tableName, Boolean createTable);

        XLHyperlink Hyperlink { get; set; }
        IXLWorksheet Worksheet { get; }

        IXLDataValidation DataValidation { get; }


        IXLCells InsertCellsAbove(int numberOfRows);
        IXLCells InsertCellsBelow(int numberOfRows);
        IXLCells InsertCellsAfter(int numberOfColumns);
        IXLCells InsertCellsBefore(int numberOfColumns);

        /// <summary>
        /// Creates a named range out of this cell. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <para>The default scope for the named range is Workbook.</para>
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLCell AddToNamed(String rangeName);

        /// <summary>
        /// Creates a named range out of this cell. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        IXLCell AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of this cell. 
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        IXLCell AddToNamed(String rangeName, XLScope scope, String comment);

        //IXLCell CopyFrom(IXLCell otherCell);

        IXLCell CopyTo(IXLCell target);

        String ValueCached { get; }

        IXLRichText RichText { get; }
        Boolean HasRichText { get; }

        Boolean IsMerged();
        Boolean IsUsed();
        Boolean IsUsed(Boolean includeFormats);
    }
}
