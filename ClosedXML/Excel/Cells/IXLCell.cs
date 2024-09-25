#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A value that is in the cell.
    /// </summary>
    public enum XLDataType
    {
        /// <summary>
        /// The value is a blank (either blank cells or the omitted optional argument of a function, e.g. <c>IF(TRUE,,)</c>.
        /// </summary>
        /// <remarks>Keep as the first, so the default values are blank.</remarks>
        Blank = 0,

        /// <summary>
        /// The value is a logical value.
        /// </summary>
        Boolean = 1,

        /// <summary>
        /// The value is a double-precision floating points number, excluding <see cref="Double.NaN"/>,
        /// <see cref="Double.PositiveInfinity"/> or <see cref="double.NegativeInfinity"/>.
        /// </summary>
        Number = 2,

        /// <summary>
        /// A text or a rich text. Can't be <c>null</c> and can be at most 32767 characters long.
        /// </summary>
        Text = 3,

        /// <summary>
        /// The value is one of <see cref="XLError"/>.
        /// </summary>
        Error = 4,

        /// <summary>
        /// The value is a <see cref="DateTime"/>, represented as a serial date time number.
        /// </summary>
        /// <remarks>
        /// Serial date time 60 is a 1900-02-29, nonexistent day kept for compatibility,
        /// but unrepresentable by <c>DateTime</c>. Don't use.
        /// </remarks>
        DateTime = 5,

        /// <summary>
        /// The value is a <see cref="TimeSpan"/>, represented in a serial date time (24 hours is 1, 36 hours is 1.5 ect.).
        /// </summary>
        TimeSpan = 6,
    }

    public enum XLTableCellType { None, Header, Data, Total }

    public interface IXLCell
    {
        /// <summary>
        /// Is this cell the <see cref="IXLWorksheet.ActiveCell">active cell of
        /// the worksheet</see>? Setting false deactivates cell only when the
        /// cell is currently active.
        /// </summary>
        Boolean Active { get; set; }

        /// <summary>Gets this cell's address, relative to the worksheet.</summary>
        /// <value>The cell's address.</value>
        IXLAddress Address { get; }

        /// <summary>
        /// Get the value of a cell without evaluation of a formula. If the cell contains
        /// a formula, it returns the last calculated value or a blank value. If the cell
        /// doesn't contain a formula, it returns same value as <see cref="Value"/>.
        /// May hold invalid value when <see cref="NeedsRecalculation"/> flag is True.
        /// </summary>
        /// <remarks>Can be useful to decrease a number of formula evaluations.</remarks>
        XLCellValue CachedValue { get; }

        /// <summary>
        /// Returns the current region. The current region is a range bounded by any combination of blank rows and blank columns
        /// </summary>
        /// <value>
        /// The current region.
        /// </value>
        IXLRange CurrentRegion { get; }

        /// <summary>
        /// Gets the type of this cell's data.
        /// </summary>
        /// <value>
        /// The type of the cell's data.
        /// </value>
        XLDataType DataType { get; }

        /// <summary>
        /// Gets or sets the cell's formula with A1 references.
        /// </summary>
        /// <remarks>Setter trims the formula and if formula starts with an <c>=</c>, it is removed.</remarks>
        /// <value>The formula with A1 references.</value>
        String FormulaA1 { get; set; }

        /// <summary>
        /// Gets or sets the cell's formula with R1C1 references.
        /// </summary>
        /// <remarks>Setter trims the formula and if formula starts with an <c>=</c>, it is removed.</remarks>
        /// <value>The formula with R1C1 references.</value>
        String FormulaR1C1 { get; set; }

        /// <summary>
        /// An indication that value of this cell is calculated by a array formula
        /// that calculates values for cells in the referenced address. Null if not part of such formula.
        /// </summary>
        IXLRangeAddress FormulaReference { get; set; }

        Boolean HasArrayFormula { get; }

        Boolean HasComment { get; }

        Boolean HasDataValidation { get; }

        Boolean HasFormula { get; }

        Boolean HasHyperlink { get; }

        Boolean HasRichText { get; }

        Boolean HasSparkline { get; }

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// Only cells with formula may return <c>true</c>, value cells always return <c>false</c>.
        /// </summary>
        Boolean NeedsRecalculation { get; }

        /// <summary>
        /// Gets or sets a value indicating whether this cell's text should be shared or not.
        /// </summary>
        /// <value>
        ///   If false the cell's text will not be shared and stored as an inline value.
        /// </value>
        Boolean ShareString { get; set; }

        IXLSparkline Sparkline { get; }

        /// <summary>
        /// Gets or sets the cell's style.
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        /// Gets or sets the cell's value.
        /// <para>
        /// Getter will return value of a cell or value of formula. Getter will evaluate a formula, if the cell
        /// <see cref="NeedsRecalculation"/>, before returning up-to-date value.
        /// </para>
        /// <para>
        /// Setter will clear a formula, if the cell contains a formula.
        /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
        /// <see cref="IXLStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
        /// </para>
        /// </summary>
        XLCellValue Value { get; set; }

        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Should the cell show phonetic (i.e. furigana) above the rich text of the cell?
        /// It shows phonetic runs in the rich text, it is not autogenerated. Default
        /// is <c>false</c>.
        /// </summary>
        Boolean ShowPhonetic { get; set; }

        IXLConditionalFormat AddConditionalFormat();

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

        /// <summary>
        /// Returns this cell as an IXLRange.
        /// </summary>
        IXLRange AsRange();

        IXLCell CellAbove();

        IXLCell CellAbove(Int32 step);

        IXLCell CellBelow();

        IXLCell CellBelow(Int32 step);

        IXLCell CellLeft();

        IXLCell CellLeft(Int32 step);

        IXLCell CellRight();

        IXLCell CellRight(Int32 step);

        /// <summary>
        /// Clears the contents of this cell.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All);

        IXLCell CopyFrom(IXLCell otherCell);

        IXLCell CopyFrom(String otherCell);

        /// <summary>
        /// Copy range content to an area of same size starting at the cell.
        /// Original content of cells is overwritten.
        /// </summary>
        /// <param name="rangeBase">Range whose content to copy.</param>
        /// <returns>This cell.</returns>
        IXLCell CopyFrom(IXLRangeBase rangeBase);

        IXLCell CopyTo(IXLCell target);

        IXLCell CopyTo(String target);

        /// <summary>
        /// Creates a new comment for the cell, replacing the existing one.
        /// </summary>
        IXLComment CreateComment();

        /// <summary>
        /// Creates a new data validation rule for the cell, replacing the existing one.
        /// </summary>
        IXLDataValidation CreateDataValidation();

        /// <summary>
        /// Creates a new hyperlink replacing the existing one.
        /// </summary>
        XLHyperlink CreateHyperlink();

        /// <summary>
        /// Replaces a value of the cell with a newly created rich text object.
        /// </summary>
        IXLRichText CreateRichText();

        /// <summary>
        /// Deletes the current cell and shifts the surrounding cells according to the shiftDeleteCells parameter.
        /// </summary>
        /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
        void Delete(XLShiftDeletedCells shiftDeleteCells);

        /// <summary>
        /// Returns the comment for the cell or create a new instance if there is no comment on the cell.
        /// </summary>
        IXLComment GetComment();

        /// <summary>
        /// Returns a data validation rule assigned to the cell, if any, or creates a new instance of data validation rule if no rule exists.
        /// </summary>
        IXLDataValidation GetDataValidation();

        /// <summary>
        /// Gets the cell's value as a Boolean.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetBoolean()</c></remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not a logical.</exception>
        Boolean GetBoolean();

        /// <summary>
        /// Gets the cell's value as a Double.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetNumber()</c></remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not a number.</exception>
        Double GetDouble();

        /// <summary>
        /// Gets the cell's value as a String.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetText()</c>. Returned value is never null.</remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not a text.</exception>
        String GetText();

        /// <summary>
        /// Gets the cell's value as a XLError.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetError()</c></remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not an error.</exception>
        XLError GetError();

        /// <summary>
        /// Gets the cell's value as a DateTime.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetDateTime()</c></remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not a DateTime.</exception>
        DateTime GetDateTime();

        /// <summary>
        /// Gets the cell's value as a TimeSpan.
        /// </summary>
        /// <remarks>Shortcut for <c>Value.GetTimeSpan()</c></remarks>
        /// <exception cref="InvalidCastException">If the value of the cell is not a TimeSpan.</exception>
        TimeSpan GetTimeSpan();

        /// <summary>
        /// Try to get cell's value converted to the T type.
        /// <para>
        /// Supported <typeparamref name="T"/> types:
        /// <list type="bullet">
        ///   <item>Boolean - uses a logic of <see cref="XLCellValue.TryConvert(out Boolean)"/></item>
        ///   <item>Number (<c>s/byte</c>, <c>u/short</c>, <c>u/int</c>, <c>u/long</c>, <c>float</c>, <c>double</c>, or <c>decimal</c>)
        ///         - uses a logic of <see cref="XLCellValue.TryConvert(out Double, System.Globalization.CultureInfo)"/> and succeeds,
        ///         if the value fits into the target type.</item>
        ///   <item>String - sets the result to a text representation of a cell value (using current culture).</item>
        ///   <item>DateTime - uses a logic of <see cref="XLCellValue.TryConvert(out DateTime)"/></item>
        ///   <item>TimeSpan - uses a logic of <see cref="XLCellValue.TryConvert(out TimeSpan, System.Globalization.CultureInfo)"/></item>
        ///   <item>XLError - if the value is of type <see cref="XLDataType.Error"/>, it will return the value.</item>
        ///   <item>Enum - tries to parse a value to a member by comparing the text of a cell value and a member name.</item>
        /// </list>
        /// </para>
        /// <para>
        /// If the <typeparamref name="T"/> is a nullable value type and the value of cell is blank or empty string, return null value.
        /// </para>
        /// <para>
        /// If the cell value can't be determined because formula function is not implemented, the method always returns <c>false</c>.
        /// </para>
        /// </summary>
        /// <typeparam name="T">The requested type into which will the value be converted.</typeparam>
        /// <param name="value">Value to store the value.</param>
        /// <returns><c>true</c> if the value was converted and the result is in the <paramref name="value"/>, <c>false</c> otherwise.</returns>
        Boolean TryGetValue<T>(out T value);

        /// <summary>
        /// <inheritdoc cref="TryGetValue{T}"/>
        /// </summary>
        /// <remarks>Conversion logic is identical with <see cref="TryGetValue{T}"/>.</remarks>
        /// <typeparam name="T">The requested type into which will the value be converted.</typeparam>
        /// <exception cref="InvalidCastException">If the value can't be converted to the type of T</exception>
        T GetValue<T>();

        /// <summary>
        /// Return cell's value represented as a string. Doesn't use cell's formatting or style.
        /// </summary>
        String GetString();

        /// <summary>
        /// Gets the cell's value formatted depending on the cell's data type and style.
        /// </summary>
        /// <param name="culture">Culture used to format the string. If <c>null</c> (default value), use current culture.</param>
        String GetFormattedString(CultureInfo culture = null);

        /// <summary>
        /// Returns a hyperlink for the cell, if any, or creates a new instance is there is no hyperlink.
        /// </summary>
        XLHyperlink GetHyperlink();

        /// <summary>
        /// Returns the value of the cell if it formatted as a rich text.
        /// </summary>
        IXLRichText GetRichText();

        IXLCells InsertCellsAbove(int numberOfRows);

        IXLCells InsertCellsAfter(int numberOfColumns);

        IXLCells InsertCellsBefore(int numberOfColumns);

        IXLCells InsertCellsBelow(int numberOfRows);

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

        /// <summary>
        /// Invalidate <see cref="CachedValue"/> so the formula will be re-evaluated next time <see cref="Value"/> is accessed.
        /// If cell does not contain formula nothing happens.
        /// </summary>
        void InvalidateFormula();

        Boolean IsEmpty();

        Boolean IsEmpty(XLCellsUsedOptions options);

        Boolean IsMerged();

        IXLRange MergedRange();

        void Select();

        IXLCell SetActive(Boolean value = true);

        [Obsolete("Use GetDataValidation to access the existing rule, or CreateDataValidation() to create a new one.")]
        IXLDataValidation SetDataValidation();

        IXLCell SetFormulaA1(String formula);

        IXLCell SetFormulaR1C1(String formula);

#nullable enable
        /// <summary>
        /// Set hyperlink of a cell. When user clicks on a cell with hyperlink,
        /// the Excel opens the target or moves cursor to the target cells in a
        /// worksheet. The text of hyperlink is a cell value, the hyperlink
        /// target and tooltip are defined by the <paramref name="hyperlink"/>
        /// parameter.
        /// </summary>
        /// <remarks>
        /// If the cell uses worksheet style, the method also sets <see cref="XLThemeColor.Hyperlink">
        /// hyperlink font color from theme</see> and the underline property.
        /// </remarks>
        /// <param name="hyperlink">The new cell hyperlink. Use <c>null</c> to
        ///   remove the hyperlink.</param>
        void SetHyperlink(XLHyperlink? hyperlink);
#nullable disable

        /// <inheritdoc cref="Value"/>
        /// <returns>This cell.</returns>
        IXLCell SetValue(XLCellValue value);

        XLTableCellType TableCellType();

        /// <summary>
        /// Returns a string that represents the current state of the cell according to the format.
        /// </summary>
        /// <param name="format">A: address, F: formula, NF: number format, BG: background color, FG: foreground color, V: formatted value</param>
        string ToString(string format);

        IXLColumn WorksheetColumn();

        IXLRow WorksheetRow();
    }
}
