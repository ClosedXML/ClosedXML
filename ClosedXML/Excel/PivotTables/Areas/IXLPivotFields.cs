// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// <para>
/// A collection of fields on <see cref="IXLPivotTable.ColumnLabels">column labels</see>,
/// <see cref="IXLPivotTable.RowLabels">row labels</see> or
/// <see cref="IXLPivotTable.ReportFilters">report filters</see> of a
/// <see cref="IXLPivotTable"/>.
/// </para>
/// </summary>
public interface IXLPivotFields : IEnumerable<IXLPivotField>
{
    IXLPivotField Add(String sourceName);

    /// <summary>
    /// Add a field to the axis labels/report filters.
    /// </summary>
    /// <param name="sourceName">Name of the field in <see cref="IXLPivotCache"/>. The value can
    ///     also be <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> for
    ///     <c>&#931;Values</c> field.</param>
    /// <param name="customName">Display name of added field. Custom name of a filed must be unique
    ///     in pivot table. Ignored for 'data' field.</param>
    /// <returns>The added field.</returns>
    /// <exception cref="ArgumentException">Field can't be added (e.g. it is already used or can't
    ///     be added to specific collection).</exception>
    IXLPivotField Add(String sourceName, String customName);

    /// <summary>
    /// Remove all fields from the axis. It also removes data of removed fields, like custom names and items.
    /// </summary>
    void Clear();

    /// <summary>
    /// Does this axis contain a field?
    /// </summary>
    /// <param name="sourceName">Name of the field in <see cref="IXLPivotCache"/>. Use
    ///     <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> for data field.</param>
    /// <returns><c>true</c> if the axis contains the field, <c>false</c> otherwise.</returns>
    Boolean Contains(String sourceName);

    /// <summary>
    /// Does this axis contain a field?
    /// </summary>
    /// <param name="pivotField">Checked pivot field.</param>
    /// <returns><c>true</c> if the axis contains the field, <c>false</c> otherwise.</returns>
    Boolean Contains(IXLPivotField pivotField);

    /// <summary>
    /// Get a field in the axis.
    /// </summary>
    /// <param name="sourceName">Name of the field in <see cref="IXLPivotCache"/> we are looking
    ///     for in the axis.</param>
    /// <returns>Found field.</returns>
    /// <exception cref="KeyNotFoundException">Axis doesn't contain field with specified name.</exception>
    IXLPivotField Get(String sourceName);

    /// <summary>
    /// Get field by index in the collection.
    /// </summary>
    /// <param name="index">Index of the field in this collection.</param>
    /// <returns>Found field.</returns>
    /// <exception cref="IndexOutOfRangeException"/>
    IXLPivotField Get(Int32 index);

    /// <summary>
    /// Get index of a field in the collection. Use the index in the <see cref="Get(Int32)"/> method.
    /// </summary>
    /// <param name="sourceName"><see cref="IXLPivotField.SourceName"/> of the field in the pivot cache.</param>
    /// <returns>Index of the field or -1 if not found.</returns>
    Int32 IndexOf(String sourceName);

    /// <summary>
    /// Get index of a field in the collection. Use the index in the <see cref="Get(Int32)"/> method.
    /// </summary>
    /// <param name="pf">Field to find. Uses <see cref="IXLPivotField.CustomName"/>.</param>
    /// <returns>Index of the field or -1 if not a member of this collection.</returns>
    Int32 IndexOf(IXLPivotField pf);

    /// <summary>
    /// Remove a field from axis. Doesn't throw, if field is not present.
    /// </summary>
    /// <param name="sourceName"><see cref="IXLPivotField.SourceName"/> of a field to remove.</param>
    void Remove(String sourceName);
}
