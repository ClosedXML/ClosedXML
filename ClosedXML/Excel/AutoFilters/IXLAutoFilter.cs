#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLFilterDynamicType { AboveAverage, BelowAverage }

    public enum XLFilterType { None, Regular, Custom, TopBottom, Dynamic }

    public enum XLTopBottomPart { Top, Bottom }

    /// <summary>
    /// <para>
    /// Autofilter can sort and filter (hide) values in a non-empty area of a sheet. Each table can
    /// have autofilter and each worksheet can have at most one range with an autofilter. First row
    /// of the area contains headers, remaining rows contain sorted and filtered data.
    /// </para>
    /// <para>
    /// Sorting of rows is done <see cref="Sort"/> method, using the passed parameters. The sort
    /// properties (<see cref="SortColumn"/> and <see cref="SortOrder"/>) are updated from
    /// properties passed to the <see cref="Sort"/> method. Sorting can be done only on values of
    /// one column.
    /// </para>
    /// <para>
    /// Autofilter can filter rows through <see cref="Reapply"/> method. The filter evaluates
    /// conditions of the autofilter and leaves visible only rows that satisfy the conditions.
    /// Rows that don't satisfy filter conditions are marked as <see cref="IXLRow.IsHidden">hidden</see>.
    /// Filter conditions can be specified for each column (accessible through <see cref="Column(string)"/>
    /// methods), e.g. <c>sheet.AutoFilter.Column(1).Top(10, XLTopBottomType.Percent)</c>
    /// creates a filter that displays only rows with values in top 10% percentile.
    /// </para>
    /// </summary>
    public interface IXLAutoFilter
    {
        /// <summary>
        /// Get rows of <see cref="Range"/> that were hidden because they didn't satisfy filter
        /// conditions during last filtering.
        /// </summary>
        /// <remarks>
        /// Visibility is automatically updated on filter change.
        /// </remarks>
        IEnumerable<IXLRangeRow> HiddenRows { get; }

        /// <summary>
        /// Is autofilter enabled? When autofilter is enabled, it shows the arrow buttons and might
        /// contain some filter that hide some rows. Disabled autofilter doesn't show arrow buttons
        /// and all rows are visible.
        /// </summary>
        Boolean IsEnabled { get; set; }

        /// <summary>
        /// Range of the autofilter. It consists of a header in first row, followed by data rows.
        /// It doesn't include totals row for tables.
        /// </summary>
        IXLRange Range { get; }

        /// <summary>
        /// What column was used during last <see cref="Sort"/>. Contains undefined value for not
        /// yet <see cref="Sorted"/> autofilter.
        /// </summary>
        Int32 SortColumn { get; }

        /// <summary>
        /// Are values in the autofilter range sorted? I.e. the values were either already loaded
        /// sorted or <see cref="Sort"/> has been called at least once.
        /// </summary>
        /// <remarks>
        /// If <c>true</c>, <see cref="SortColumn"/> and <see cref="SortOrder"/> contain valid values.
        /// </remarks>
        Boolean Sorted { get; }

        /// <summary>
        /// What sorting order was used during last <see cref="Sort"/>. Contains undefined value
        /// for not yet <see cref="Sorted"/> autofilter.
        /// </summary>
        XLSortOrder SortOrder { get; }

        /// <summary>
        /// Get rows of <see cref="Range"/> that are visible because they satisfied filter
        /// conditions during last filtering.
        /// </summary>
        /// <remarks>
        /// Visibility is not updated on filter change.
        /// </remarks>
        IEnumerable<IXLRangeRow> VisibleRows { get; }

        /// <summary>
        /// Disable autofilter, remove all filters and unhide all rows of the <see cref="Range"/>.
        /// </summary>
        IXLAutoFilter Clear();

        /// <summary>
        /// Get filter configuration for a column.
        /// </summary>
        /// <param name="columnLetter">
        /// Column letter that determines number in the range, from <em>A</em> as the first column
        /// of a <see cref="Range"/>.
        /// </param>
        /// <returns>Filter configuration for the column.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Invalid column.</exception>
        IXLFilterColumn Column(String columnLetter);

        /// <summary>
        /// Get filter configuration for a column.
        /// </summary>
        /// <param name="columnNumber">Column number in the range, from 1 as the first column of a <see cref="Range"/>.</param>
        /// <returns>Filter configuration for the column.</returns>
        IXLFilterColumn Column(Int32 columnNumber);

        /// <summary>
        /// Apply autofilter filters to the range and show every row that satisfies the conditions
        /// and hide the ones that don't satisfy conditions.
        /// </summary>
        /// <remarks>
        /// Filter is generally automatically applied on a filter change. This method could be
        /// called after a cell value change or row deletion.
        /// </remarks>
        IXLAutoFilter Reapply();

        /// <summary>
        /// Sort rows of the range using data of one column.
        /// </summary>
        /// <remarks>
        /// This method sets <see cref="Sorted"/>, <see cref="SortColumn"/> and <see cref="SortOrder"/> properties.
        /// </remarks>
        /// <param name="columnToSortBy">
        /// Column number in the range, from 1 to width of the <see cref="Range"/>.
        /// </param>
        /// <param name="sortOrder">Should rows be sorted in ascending or descending order?</param>
        /// <param name="matchCase">
        /// Should <see cref="XLDataType.Text"/> values on the column be matched case sensitive.
        /// </param>
        /// <param name="ignoreBlanks">
        /// <c>true</c> - rows with blank value in the column will always at the end, regardless of
        /// sorting order. <c>false</c> - blank will be treated as empty string and sorted
        /// accordingly.
        /// </param>
        IXLAutoFilter Sort(Int32 columnToSortBy = 1, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
    }
}
