#############################
Migration from 0.105 to 0.106
#############################

**************
Number formats
**************

Format is always there
----------------------

Number format ```IXLNumberFormatBase.Format``` previously returned an empty
string for predefined formats. It now  returns actually used predefined format
code instead.

Number format id
----------------

The number format setters (```IXLNumberFormatBase.NumberFormatId```,
```IXLNumberFormat.SetNumberFormatId(int)``` and ```IXLPivotValueFormat.SetNumberFormatId(int)```)
now throw an ```ArgumentOutOfRangeException``` when supplied number format id is
not a predefined format id from ```XLPredefinedFormat```.

************
IXLAlignment
************

The ``IXLAlignment.TextRotation`` now throws ``ArgumentOutOfRangeException`` on invalid rotation
instead of original ``ArgumentException``.

*************
IXLWorksheets
*************

The method ``IXLWorksheets.Worksheet(string sheetName)`` now throws ``KeyNotFoundException`` when
sheet is not found, instead of original ``ArgumentException``.

Generally speaking, get-by-name methods of all unique-name collections (e.g., worksheets, styles,
table fields, etc.) should throw ``KeyNotFoundException`` when an item is not found. This change
aligns the behavior of the ``IXLWorksheets`` with the rest of the API.

*************
Defined names
*************

A property setter ``IXLDefinedName.RefersTo`` now throws an ``ArgumentException``
when trying to set an empty/whitespace-only value.

************
IXLRangeBase
************

An obsolete method ``IXLRangeBase.SetDataValidation`` has been removed. Use ``GetDataValidation()``
to access the existing rule, or ``CreateDataValidation()`` to create a new one.

*******
IXLCell
*******

An obsolete method ``IXLRangeBase.SetDataValidation`` has been removed. Use ``GetDataValidation()``
to access the existing rule, or ``CreateDataValidation()`` to create a new one.
