############################
Migration from 0.97 to 0.100
############################

******************************
Strongly typed value of a cell
******************************

The key breaking change is that the ``IXLCell.Value`` is no longer untyped
``Object`` backed by a string interpreted through a ``XLDataType``, but
instead is strongly typed readonly structure ``XLCellValue`` that can represent
any value of a cell. Type and value are now intrinsically linked together and
it is not possible to change data type without changing value.

Cell value (through ``XLCellValue``) can now be an ``XLError``, either literal
or as a result of formula calculation.

Strongly typed cell value 
=========================

``IXLCell.Value`` and ``IXLCell.CachedValue`` are now of type ``XLCellValue``.
All possible values of a cell (blank, logical, number, text, error) can be
converted to ``XLCellValue`` through implicit casting operators.

Due to change of the ``Value`` setter, it is no longer possible use setter to

* Set the value by the ``IRichText``. Use ``IXLCell.GetRichText().CopyFrom(IRichText)`` instead.
* Set the value by the ``DateTimeOffset``. Use implicit ``XLCellValue`` cast operator from ``DateTimeOffset.Date`` instead.
* Set the value by the ``Guid``. Use implicit ``XLCellValue`` cast operator from ``Guid.ToString()`` instead.
* Inserting data by setting a value of type ``IEnumerable``. Use either ``IXLCell.InsertData(IEnumerable)`` or ``IXLCell.InsertData<T>(IEnumerable<T>)``.
* Copy data by setting a value of type ``IXLRangeBase``. Use ``IXLCell.CopyFrom(IXLRangeBase)``
* Set a value to an object of any type. It originally took an object and used its ``ToString()`` method to convert the object. Call the ``ToString()`` directly
  in the code before setting the value to a string.
* It is no longer possible set a value ``NaN`` or ``Infinity``.

``SetDataType`` methods removed
===============================

Method ``SetDataType`` has been removed from all interfaces (``IXLCell``,
``IXLColumn``, ``IXLColumns``, ``IXLRange`` ...). There is no replacement, if you
need to reinterpret existing data, do it in application code and set a new value
with a specific type.

Evaluate methods
================

Evaluation methods ``IXLWorkbook.Evaluate(String)``, ``XLWorkbook.EvaluateExpr(String)``
and ``IXLWorksheet.Evaluate(String, String)`` don't return ``Object``, but
``XLCellValue``.

Bulk data insert
================

Previously, it was possible to insert data into a worksheet by calling
a ``IXLCell.Value`` setter with a value of ``IEnumerable``. ``IXLCell.Value``
no longer accepts object, use ``IXLCell.InsertData`` methods instead.

Bulk copy cell values
=====================

Previously, it was possible to copy data from a range of cells to cells
starting at cell by calling a ``IXLCell.Value`` setter with a value of
``IXLRangeBase``. ``IXLCell.Value`` no longer accepts ``IXLRangeBase``,
use ``IXLCell.CopyFrom`` methods instead.

Rich text connected to cell value
=================================

Previously, it was possible to set a rich text to a cell by calling
a ``IXLCell.Value`` setter with a value of ``IXLRichText``. ``IXLRichText``
is now connected to the cell, changing a value of the rich text also changes
value of the cell the rich text belongs to.

As a conseqence, rich text can longer be copied around from one cell
to another. If you need to copy a rich text from one cell to another, use
``IXLRichText.CopyFrom`` method.

.. code-block:: csharp

   var cell = ws.Cell(1,1);
   var richText = cell.GetRichText();

   richText.AddText("Hello").SetFontSize(15);
   Assert.AreEqual("Hello", cell.Value);

   richText.AddText("World").SetFontSize(20);
   Assert.AreEqual("HelloWorld", cell.Value);


Copy cell value
===============

Previously, it was possible to use ``IXLCell.Value`` setter to copy a different
cell to a cell. The main benefit in comparison of just copying the value was
copying of conditional formatting of original cell. Conditional formatting is
still copied for ``IXLCell.CopyFrom``, so use ``IXLCell.AsRange()`` method as
an intermediate step during replacement.

.. code-block:: csharp

   var sourceCell = ws.Cell(1, 1);
   var targetCell = ws.Cell(2, 1);
   targetCell.CopyFrom(sourceCell.AsRange());


Data type detected removed
==========================

Edge double values like ``Double.NaN``, ``Double.PositiveInfinity``,
``Double.NegativeInfinity`` can't be excel cell value. Previously, such values
were converted to string, leading to "saving number, getting text" situations.
``XLCellValue`` now throws an ``ArgumentException`` on initialization from such
values.

ClosedXML also previously sometimes incorrectly detected string as a date time
(e.g. for *"Z12.31"* interpreted as *2022-12-31*). Whole detection has been
removed, developer is now in control of the type in a cell through
``XLCellValue``.

TryGetValue changes
===================

Previously, it was possible to retrive a ``IXLRichText`` or ``XLHyperlink``
component of a cell through ``IXLCell.TryGetValue``. That is no longer
possible, use ``IXLCell.GetRichText()`` or ``IXLCell.GetHyperlink()``.

DateTime pre-1900
=================

Previously, dates before 1900-01-01 were converted to text. That no longer
happens, it is possible to set value to any ``DateTime`` value. The cell type
``XLDataType.DateTime`` is mostly masquarade above serial date time, values
before 1900 are displayed as *######*, but are still a serial date time values.

XLClearOptions.DataType removed
===============================

The enum member ``XLClearOptions.DataType`` has been removed. It makes no
semantic sense, if you need to clear data type, you must set a new value. Use
``IXLRangeBase.SetValue`` or ``IXLCell.SetValue`` instead.

Cast errors throw InvalidCastException
======================================

Previously, methods to get a value of a cell used to the throw
``FormatException``, instead they now throw ``InvalidCastException`` (+ they
are now mostly shortcut to ``XLCellValue`` methods).

* ``IXLCell.GetBoolean()``
* ``IXLCell.GetDouble()``
* ``IXLCell.GetDateTime()``
* ``IXLCell.GetTimeSpan()``

Method ``IXLCell.GetValue<T>()`` now also throws ``InvalidCastException``
instead of ``FormatException``.

IXLWorksheet.Search
===================

``IXLWorksheet.Search`` searches in the value text representation, not
formatted string. That is consistent with Excel search behavior.

An example for a number **12345.7** for a culture with a decimal separator *,*

* Formatting (``IXLCell.GetFormattedString()``) adds thousand separator and
  the value is formatted as ``12 345,7`` in a cell
* In the formula bar, the value is represented as a ``12345,7`` (text
  representation)
* Searching for a string ``2345,7`` will find the value, because it is
  a substring of text representation

Pivot table values use XLCellValue
==================================

Previously, the predicate of ``IXLPivotValueStyleFormat.AndWith`` (used to
specify which values to apply style to) has an ``Object`` as a parameter of
a predicate. It now has parameter of type ``XLCellValue``.

It also applies to several other API:

* ``IXLPivotField.SelectedValues``
* ``IXLPivotField.AddSelectedValue``
* ``IXLPivotField.AddSelectedValues``

*****************
CalcEngine errors
*****************

Previously, if an error happened during formula evaluation (e.g. division by
``=1/0`` `#DIV/0!`) have thrown an exception for the error derived from
``CalcEngineException``. Errors have been incorporated to CalcEngine and are
now a valid value that can be stored in a cell or it can be a result of formula
evaluation.

Errors are represented by an ``XLError`` enum. ``CalcEngineException`` and
derived exception have been removed.

.. code-block:: csharp

   // Errors are now valid return value. CalcEngine no longer throws exceptions
   Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("1/0"));


Previously, if formula contained a standard unimplemented function,
``NameNotRecognizedException`` was thrown during parsing. Instead CalcEngine
will now return ``XLError.NameNotRecognized`` error.

.. code-block:: csharp

   var wb = new XLWorkbook();
   var ws = wb.AddWorksheet();
   var cell = ws.Cell(1,1);
   cell.FormulaA1 = "RTD(\"stockprice.rtd\", \"NASD\", \"MSFT\")";
   var value = cell.Value; // Used to throw NameNotRecognizedException
   Assert.AreEqual(XLError.NameNotRecognized, value.GetError());


This causes a differences, if ClosedXML saves formula values (by default it
doesn't, but can be enabled by ``SaveOptions.EvaluateFormulasBeforeSaving``).
The original behavior kept the values blank for cells with formulas containing
unimplemented functions, new behavior will set values of cells to ``#NAME?``
User won't see a difference, because Excel recalculates values on load (this
is the default calculate mode for workbooks). If the workbook has a different
mode (e.g. ``XLWorkbook.CalculateMode = XLCalculateMode.Manual``), user might
see the ``#NAME?`` values instead of blanks in some formulas.

************************************
XLError enum moved and order changed
************************************

Enum XLError has been moved from ``ClosedXML.Excel.CalcEngine`` namespace
to ``ClosedXML.Excel`` namespace. XLError's members have been reordered, so
the order is same as values returned by ERROR.TYPE function (the values
are actually used sometimes during sorting).

****************
Value formatting
****************

Previously, ``IXLCell.GetFormattedString()`` formatted logical values ``true``/``false`` to a string *True*/*False*. It now formats them to Excel compliant *TRUE*/*FALSE*.

***********************
Pivot table value field
***********************

Methods for manipulating the ``IXLPivotValues`` now use the custom name of
a pivot value fields, not source names. Source name is roughly name of
a column in the source table while custom name is a name of a field in
the pivot table. There can be multiple values for a single source column
(e.g. average value and minimal value).

Methods for manipulating the ``IXLPivotFields`` still use source names.

***********************
XLEventTracking removed
***********************

ClosedXML used to track various events and call registered event handlers.
That functionality was removed long ago and now even enum
``XLEventTracking``, ``LoadOptions.EventTracking`` property
and ``XLWorkbook`` constructors that accepted the enum were removed.

To migrate the code, just remove the ``XLEventTracking`` argument from
the constructor and remove setters of ``LoadOptions.EventTracking``.