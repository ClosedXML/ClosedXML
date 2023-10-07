using System;
using System.Collections.Generic;
using ClosedXML.Excel.Exceptions;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A cache of pivot data - essentially a collection of fields and their values that can be
    /// displayed by a <see cref="IXLPivotTable"/>. Data for the cache are retrieved from
    /// an area (a table or a range). The pivot cache data are <strong>cached</strong>, i.e.
    /// the data in the source are not immediately updated once the data in a worksheet change.
    /// </summary>
    public interface IXLPivotCache
    {
        /// <summary>
        /// Get names of all fields in the source, in left to right order. Every field name is unique.
        /// </summary>
        /// <remarks>
        /// The field names are case insensitive. The field names of the cached
        /// source might differ from actual names of the columns
        /// in the data cells.
        /// </remarks>
        IReadOnlyList<String> FieldNames { get; }

        /// <summary>
        /// Gets the number of unused items in shared items to allow before discarding unused items.
        /// </summary>
        /// <remarks>
        /// Shared items are distinct values of a source field values. Updating them can be expensive
        /// and this controls, when should the cache be updated. Application-dependent attribute.
        /// </remarks>
        /// <value>Default value is <see cref="XLItemsToRetain.Automatic"/>.</value>
        XLItemsToRetain ItemsToRetainPerField { get; set; }

        /// <summary>
        /// Will Excel refresh the cache when it opens the workbook.
        /// </summary>
        /// <value>Default value is <c>false</c>.</value>
        Boolean RefreshDataOnOpen { get; set; }

        /// <summary>
        /// Should the cached values of the pivot source be saved into the workbook file?
        /// If source data are not saved, they will have to be refreshed from the source
        /// reference which might cause a change in the table values.
        /// </summary>
        /// <value>Default value is <c>true</c>.</value>
        Boolean SaveSourceData { get; set; }

        /// <summary>
        /// Refresh data in the pivot source from the source reference data.
        /// </summary>
        /// <exception cref="InvalidReferenceException">The data source for the pivot table can't be found.</exception>
        IXLPivotCache Refresh();

        /// <inheritdoc cref="ItemsToRetainPerField"/>
        IXLPivotCache SetItemsToRetainPerField(XLItemsToRetain value);

        /// <inheritdoc cref="RefreshDataOnOpen"/>
        /// <remarks>Sets the value to <c>true</c>.</remarks>
        IXLPivotCache SetRefreshDataOnOpen();

        /// <inheritdoc cref="RefreshDataOnOpen"/>
        IXLPivotCache SetRefreshDataOnOpen(Boolean value);

        /// <inheritdoc cref="SaveSourceData"/>
        /// <remarks>Sets the value to <c>true</c>.</remarks>
        IXLPivotCache SetSaveSourceData();

        /// <inheritdoc cref="SaveSourceData"/>
        IXLPivotCache SetSaveSourceData(Boolean value);
    }
}
