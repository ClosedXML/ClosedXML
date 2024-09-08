#############################
Migration from 0.103 to 0.104
#############################

************************
Minimal required version
************************

OpenXML SDK dependency has been upgraded to 3.0.

Minimal required version for .NET Framework has been increased from net461 to
net462. The net461 didn't support netstandard 2.0 properly and OpenXML SDK 3.0
requires net462.

******************
Throw on not found
******************

Several API used to return ``null``, when the searched element wasn't found.
They now throw an ``ArgumentException`` instead.

If you need to avoid the exception, use methods these methods like
``IXLNamedRanges.TryGetValue(string rangeName, out IXLNamedRange range)`` or
* ``IXLNamedRanges.Contains(string rangeName)``.

``IXLWorksheet.Cell(string)``
#############################

``IXLWorksheet.Cell(string cellAddressInRange)`` used to return ``null`` when
the ``cellAddressInRange`` wasn't A1 address or workbook scoped named range.

It now throws ``ArgumentException`` instead.

``IXLWorksheet.NamedRange(string)``
###################################

``IXLWorksheet.NamedRange(string rangeName)`` used to return ``null`` when
the ``rangeName`` wasn't found.

It now throws ``ArgumentException`` instead.

``IXLWorksheet.Range(string)``
##############################

``IXLWorksheet.Range(string rangeAddress)`` used to return ``null`` when
the ``rangeAddress`` wasn't A1 address or named range.

It now throws ``ArgumentException`` instead.

``IXLNamedRanges.NamedRange(string)``
###################################

``IXLNamedRanges.NamedRange(string rangeName)`` used to return ``null`` when
the ``nameRange`` wasn't found.

It now throws ``ArgumentException`` instead.

**************
IXLPivotTables
**************

``IXLPivotTables.Add`` method always first looks for a table with same area as
passed range and if one is found, the table itself is used as a source for the
pivot cache.

.. code-block:: csharp

   // The workbook already contains a table A1:B3
   var range = ws.Range("A1:A3");
   // Although we passed a range and there isn't any pivot cache, the added
   // pivot cache uses the table as source, not the range.
   var pivot = ws.PivotTables.Add("pivot table", ws.Cell("A1"), range);

Generally, this change doesn't matter, unless the table changes sizes.

*********************
Pivot table subtotals
*********************

``XLSubtotalFunction.None`` has been removed and ``XLSubtotalFunction.Maximum``
and ``XLSubtotalFunction.Minimum`` have changed order in the enum declaration.

``IXLPivotField.Subtotals`` is no longer modifiable list, but only read only
list. At this time, there is no function to remove subtotal function for a field.

***************************
Pivot table fileter styling
***************************

Ability to style filter area in pivot tables has been removed. The API is still
there under ```pivotTable.ReportFilters.Get("filter field).StyleFormats```, but
it throws ```NotImplementedException```.

It will get re-implemented in a later versions.


*******
Sorting
*******

Sorting algorithm has been modified, so it matches Excel. It now sorts values
first by type (number, text, logical, error, blank), then by value. Blanks are
always last, regardless of sorting order (unless ``ignoreBlanks`` is set to
``false``).

* ``IXLSortElement`` properties no longer have setters.
* An unused enum ``XLSortOrientation`` has been deleted.

**********
Page setup
**********

``IXLPageSetup.FirstPageNumber`` and ``IXLPageSetup.SetFirstPageNumber(int)``
now use ``int`` type instead of ``uint``. First page number can be negative and
``int`` is thus better (``-3`` instead of ``4294967293``).

**********
AutoFilter
**********

``IXLFilterColumn.AddFilter`` and ``IXLFilteredColumn.AddFilter`` method
parameter type was changed from a generic ``T : IComparable<T>`` to ``XLCellValue``.
Semantic of method was also updated to reflect how Excel actually filter column
values.

Removed setters for autofilter configuration, the setters were given access to
internal state and the only acceptable way to set filters is through
``IXLFilterColumn`` methods. 

Following methods were removed.

* ``IXLAutoFilter.Range`` setter.
* ``IXLAutoFilter.SortColumn`` setter.
* ``IXLAutoFilter.Sorted`` setter.
* ``IXLAutoFilter.SortOrder`` setter.
* ``IXLFilterColumn.FilterType`` setter.
* ``IXLFilterColumn.SetFilterType(XLFilterType value)``
* ``IXLFilterColumn.TopBottomValue`` setter.
* ``IXLFilterColumn.SetTopBottomValue(Int32 value)``
* ``IXLFilterColumn.TopBottomType`` setter.
* ``IXLFilterColumn.SetTopBottomType(XLTopBottomType value)``
* ``IXLFilterColumn.TopBottomPart`` setter.
* ``IXLFilterColumn.SetTopBottomPart(XLTopBottomPart value)``
* ``IXLFilterColumn.DynamicType`` setter.
* ``IXLFilterColumn.SetDynamicType(XLFilterDynamicType value)``
* ``IXLFilterColumn.DynamicValue`` setter.
* ``IXLFilterColumn.SetDynamicValue(Double value)``

Added a new type of filter (``XLFilterType.None``) that is used when autofilter
doesn't have any filter.

The filter type ``XLFilterType.DateTimeGrouping`` has been removed. It was an
artifical type, the actual filter type is ``XLFilterType.Regular``. The removal
allows to use regular and date time grouping in one filter column at once.

The interface ``IXLDateTimeGroupFilteredColumn`` has been merged into
``IXLFilteredColumn``. That allows to specify both date time group and values
for regular filter in same fluent API.

Methods that add/set filters now have an ``bool`` parameter ``reapply``. By
default, it is set to ``true``. The parameter determines if the method should
immediately reapplied modified filters to the autofilter. This makes it possile
to configure several filters and only then call ``IXLAutoFilter.Reapply()``.

Method ``IXLFilterColumn.Top`` and ``IXLFilterColumn.Bottom`` now throw an
``ArgumentOutOfRangeException`` when passed item count or percentage is not
between 1 and 500.

Method ``IXLFilterColumn.Clear`` now has a new parameter ``reapply`` (set by default to true to
match the rest of methods) that determines if filters should be reapplied after cleaing column
filter. Originally, there wasn't any parameter and clearing didn't reapply filters.

*******
IXLCell
*******

``IXLCell.GetFormattedString(CultureInfo)`` now has an optional argument for a
culture. By default, it uses current culture in all cases (was inconsistent),
but culture can be explicitely specified.

********
IXLStyle
********

``IXLStyle.Equals`` method (it's implementor) now compares equality purely by style properties.
Originally, it also checked the container equality and thus were rarely equal. Because styles are
internally immutable, the ``IXLStyle`` object must hold a reference to object that contains the
immutable style in a property (e.g. ``IXLCell`` or ``IXLRow``) so it can change it and that
reference is called container. The end result is that two IXLStyle objects should be equal when all
their style properties are equal.

*************
Defined names
*************

``IXLWorksheet.NamedRange(string)`` throws ``KeyNotFoundException`` instead of
``ArgumentOutOfRangeException`` when defined name is not found.

Names of interfaces has been changed to better reflect semantic meaning, i.e. defined name. Defined
name can refer to a range, constant, cell, function, lambda and others. *named range* is very
non-descript type name.

* ``IXLNamedRange`` -> ``IXLDefinedName``
* ``IXLNamedRanges`` -> ``IXLDefinedNames``

Various properties/names containing ``*NamedRange*`` have been renamed to ``*DefinedName*`` and
marked with an ``[Obsolete]`` attribute pointing to a new name.

The source of truth in a defined name is ``IXLDefinedName.RefersTo``, it used to be
``IXLNamedRange.Ranges``. The formula in defined name is now parsed and validated when it is being
set, so it might throw an exception. The redundant equal sign (``=``) is now also removed from
formula in the setter.

``IXLDefinedName.Clear()`` has been removed. It makes no sense to have an operation that turns
defined range to a non-valid (=empty) formula.

Methods to modify the defined name by adding/removing ranges from a list of ranges in formula have
been removed. Methods only makes sense when defined name represents a union of ranges, but that is
not always the case. If you need to modify the name, create a new one formula of range unions and
set through ``IXLDefinedName.SetRefersTo(string)``. List of removed methods:

* ``IXLDefinedName.Add(IXLRange range)``
* ``IXLDefinedName.Add(IXLRanges ranges)``
* ``IXLDefinedName.Add(XLWorkbook workbook, String rangeAddress);
* ``IXLDefinedName.Remove(String rangeAddress)``
* ``IXLDefinedName.Remove(IXLRange range)``
* ``IXLDefinedName.Remove(IXLRanges ranges)``

``IXLDefinedName.Copyto(IXLWorksheet targetSheet)`` now throws an exception when copied name is not
sheet-scoped and it copies ranges and tables referencing the original sheet, if found in the new
sheet.

*********
Worksheet
*********

Changing a worksheet name through ``IXLWorksheet.Name`` setter now also changes names in formulas
and defined names that use the original sheet name.
