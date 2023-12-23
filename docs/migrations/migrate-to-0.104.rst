#############################
Migration from 0.103 to 0.104
#############################

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

*******
Sorting
*******

* ``IXLSortElement`` properties no longer have setters.
* An unused enum ``XLSortOrientation`` has been deleted.