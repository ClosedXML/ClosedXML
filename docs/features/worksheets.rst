*********************
Organizing Worksheets
*********************

It is possible to use ClosedXML to organize the worksheets in a workbook.

Adding a worksheet
------------------

There are several overloads of the method `AddWorksheet` that take a name and a position (first position is 1). If an argument is missing, ClosedXML will use a default behavior.
For name, a new name `Sheet{number}` will be used. If a position is missing, sheet will be added as the last worksheet of a workbook.

.. code-block:: csharp

   // Add a 'Sheet1' as the last worksheet, in the case of a new workbook as a first sheet
   wb.AddWorksheet();

   // Add a worksheet with a name 'Export' as the last sheet of a workbook
   wb.AddWorksheet("Export");

   // Add a worksheet Import at position 2, moving all other sheets to the right
   // 'Export' will be in the last position, the end result will be 'Sheet1', 'Import', 'Export'
   wb.AddWorksheet("Import", 2);

Methods `XLWorkbook.Worksheets.Add` are behaving the same way.

Removing a worksheet
--------------------
Worksheet can be removed using a position or a sheet name.

.. code-block:: csharp

   wb.Worksheets.Delete('Export');

   wb.Worksheets.Delete(2);

Moving worksheets
-----------------
It is also possible to rearrange the order of the worksheets in the workbook. The worksheets will move from original position and will be in a new position, moving other sheets accordingly.

.. code-block:: csharp

   wb.AddWorksheet("Sheet1");
   wb.AddWorksheet("Sheet2");
   wb.AddWorksheet("Sheet3");
   wb.AddWorksheet("Sheet4");

   // The end result will be Sheet2, Sheet3, Sheet1, Sheet4
   wb.Worksheet("Sheet1").Position = 3;
