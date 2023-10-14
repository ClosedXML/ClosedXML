#############################
Migration from 0.103 to 0.104
#############################

************
IXLWorksheet
************

``IXLWorksheet.Cell(string cellAddressInRange)`` used to return ``null`` when
the ``cellAddressInRange`` wasn't A1 address or workbook scoped named range.

It now throws ``ArgumentOutOfRangeException`` instead.

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
