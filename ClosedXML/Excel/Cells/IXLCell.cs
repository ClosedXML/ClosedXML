using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ClosedXML.Excel
{
    public enum XLDataType { Text, Number, Boolean, DateTime, TimeSpan }

    public enum XLTableCellType { None, Header, Data, Total }

    public interface IXLCell
    {
        /// <summary>
        /// Gets or sets the cell's value. To get or set a strongly typed value, use the GetValue&lt;T&gt; and SetValue methods.
        /// <para>ClosedXML will try to detect the data type through parsing. If it can't then the value will be left as a string.</para>
        /// <para>If the object is an IEnumerable, ClosedXML will copy the collection's data into a table starting from this cell.</para>
        /// <para>If the object is a range, ClosedXML will copy the range starting from this cell.</para>
        /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
        /// </summary>
        /// <value>
        /// The object containing the value(s) to set.
        /// </value>
        Object Value { get; set; }

        /// <summary>Gets this cell's address, relative to the worksheet.</summary>
        /// <value>The cell's address.</value>
        IXLAddress Address { get; }

        /// <summary>
        /// Returns the current region. The current region is a range bounded by any combination of blank rows and blank columns
        /// </summary>
        /// <value>
        /// The current region.
        /// </value>
        IXLRange CurrentRegion { get; }

        /// <summary>
        /// Gets or sets the type of this cell's data.
        /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
        /// </summary>
        /// <value>
        /// The type of the cell's data.
        /// </value>
        /// <exception cref="ArgumentException"></exception>
        XLDataType DataType { get; set; }

        /// <summary>
        /// Sets the type of this cell's data.
        /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
        /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
        /// </summary>
        /// <param name="dataType">Type of the data.</param>
        /// <returns></returns>
        IXLCell SetDataType(XLDataType dataType);

        /// <summary>
        /// Sets the cell's value.
        /// <para>If the object is an IEnumerable ClosedXML will copy the collection's data into a table starting from this cell.</para>
        /// <para>If the object is a range ClosedXML will copy the range starting from this cell.</para>
        /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
        /// <para>ClosedXML will try to translate it to the corresponding type, if it can't then the value will be left as a string.</para>
        /// </summary>
        /// <value>
        /// The object containing the value(s) to set.
        /// </value>
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

        XLHyperlink GetHyperlink();

        Boolean TryGetValue<T>(out T value);

        Boolean HasHyperlink { get; }

        /// <summary>
        /// Clears the contents of this cell.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All);

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

        IXLCell SetFormulaA1(String formula);

        /// <summary>
        /// Gets or sets the cell's formula with R1C1 references.
        /// </summary>
        /// <value>The formula with R1C1 references.</value>
        String FormulaR1C1 { get; set; }

        IXLCell SetFormulaR1C1(String formula);

        /// <summary>
        /// Returns this cell as an IXLRange.
        /// </summary>
        IXLRange AsRange();

        /// <summary>
        /// Gets or sets the cell's style.
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this cell's text should be shared or not.
        /// </summary>
        /// <value>
        ///   If false the cell's text will not be shared and stored as an inline value.
        /// </value>
        Boolean ShareString { get; set; }

        /// <summary>
        /// Inserts the IEnumerable data elements and returns the range it occupies.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        IXLRange InsertData(IEnumerable data);

        /// <summary>
        /// Inserts the IEnumerable data elements and returns the range it occupies.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
        /// <returns></returns>
        IXLRange InsertData(IEnumerable data, Boolean transpose);

        /// <summary>
        /// Inserts the data of a data table.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <returns>The range occupied by the inserted data</returns>
        IXLRange InsertData(DataTable dataTable);

        /// <summary>
        /// Inserts the IEnumerable data elements as a table and returns it.
        /// <para>The new table will receive a generic name: Table#</para>
        /// </summary>
        /// <param name="data">The table data.</param>
        IXLTable InsertTable<T>(IEnumerable<T> data);

        /// <summary>
        /// Inserts the IEnumerable data elements as a table and returns it.
        /// <para>The new table will receive a generic name: Table#</para>
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="createTable">
        /// if set to <c>true</c> it will create an Excel table.
        /// <para>if set to <c>false</c> the table will be created in memory.</para>
        /// </param>
        IXLTable InsertTable<T>(IEnumerable<T> data, Boolean createTable);

        /// <summary>
        /// Creates an Excel table from the given IEnumerable data elements.
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="tableName">Name of the table.</param>
        IXLTable InsertTable<T>(IEnumerable<T> data, String tableName);

        /// <summary>
        /// Inserts the IEnumerable data elements as a table and returns it.
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="createTable">
        /// if set to <c>true</c> it will create an Excel table.
        /// <para>if set to <c>false</c> the table will be created in memory.</para>
        /// </param>
        IXLTable InsertTable<T>(IEnumerable<T> data, String tableName, Boolean createTable);

        /// <summary>
        /// Inserts the DataTable data elements as a table and returns it.
        /// <para>The new table will receive a generic name: Table#</para>
        /// </summary>
        /// <param name="data">The table data.</param>
        IXLTable InsertTable(DataTable data);

        /// <summary>
        /// Inserts the DataTable data elements as a table and returns it.
        /// <para>The new table will receive a generic name: Table#</para>
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="createTable">
        /// if set to <c>true</c> it will create an Excel table.
        /// <para>if set to <c>false</c> the table will be created in memory.</para>
        /// </param>
        IXLTable InsertTable(DataTable data, Boolean createTable);

        /// <summary>
        /// Creates an Excel table from the given DataTable data elements.
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="tableName">Name of the table.</param>
        IXLTable InsertTable(DataTable data, String tableName);

        /// <summary>
        /// Inserts the DataTable data elements as a table and returns it.
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="createTable">
        /// if set to <c>true</c> it will create an Excel table.
        /// <para>if set to <c>false</c> the table will be created in memory.</para>
        /// </param>
        IXLTable InsertTable(DataTable data, String tableName, Boolean createTable);

        XLTableCellType TableCellType();

        XLHyperlink Hyperlink { get; set; }
        IXLWorksheet Worksheet { get; }

        IXLDataValidation DataValidation { get; }
        IXLDataValidation NewDataValidation { get; }

        IXLDataValidation SetDataValidation();

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
        /// </summary>
        IXLCell AddToNamed(String rangeName, XLScope scope);

        /// <summary>
        /// Creates a named range out of this cell.
        /// <para>If the named range exists, it will add this range to that named range.</para>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="scope">The scope for the named range.</param>
        /// <param name="comment">The comments for the named range.</param>
        /// </summary>
        IXLCell AddToNamed(String rangeName, XLScope scope, String comment);

        IXLCell CopyFrom(IXLCell otherCell);

        IXLCell CopyFrom(String otherCell);

        IXLCell CopyTo(IXLCell target);

        IXLCell CopyTo(String target);

        /// <summary>
        /// Textual representation of cell calculated value (as it is saved to a workbook or read from it)
        /// </summary>
        [Obsolete("Use CachedValue instead")]
        String ValueCached { get; }

        /// <summary>
        /// Calculated value of cell formula. Is used for decreasing number of computations perfromed.
        /// May hold invalid value when <see cref="NeedsRecalculation"/> flag is True.
        /// </summary>
        Object CachedValue { get; }

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// </summary>
        Boolean NeedsRecalculation { get; }

        /// <summary>
        /// Invalidate <see cref="CachedValue"/> so the formula will be re-evaluated next time <see cref="Value"/> is accessed.
        /// If cell does not contain formula nothing happens.
        /// </summary>
        void InvalidateFormula();

        IXLRichText RichText { get; }
        Boolean HasRichText { get; }
        IXLComment Comment { get; }
        Boolean HasComment { get; }

        Boolean IsMerged();

        IXLRange MergedRange();

        Boolean IsEmpty();

        Boolean IsEmpty(Boolean includeFormats);

        IXLCell CellAbove();

        IXLCell CellAbove(Int32 step);

        IXLCell CellBelow();

        IXLCell CellBelow(Int32 step);

        IXLCell CellLeft();

        IXLCell CellLeft(Int32 step);

        IXLCell CellRight();

        IXLCell CellRight(Int32 step);

        IXLColumn WorksheetColumn();

        IXLRow WorksheetRow();

        Boolean HasDataValidation { get; }

        IXLConditionalFormat AddConditionalFormat();

        void Select();

        Boolean Active { get; set; }

        IXLCell SetActive(Boolean value = true);

        Boolean HasFormula { get; }
        Boolean HasArrayFormula { get; }

        IXLRangeAddress FormulaReference { get; set; }
    }
}
