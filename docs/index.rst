ClosedXML
*********

ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and
user-friendly interface to dealing with the underlying OpenXML API.

Install the library through .NET CLI

.. code-block:: batch

   C:\source> dotnet add package ClosedXML

ClosedXML allows you to create Excel files without the Excel application. The typical example is creating Excel reports on a web server.

.. code-block:: csharp

   using var workbook = new XLWorkbook();
   var worksheet = workbook.AddWorksheet("Sample Sheet");
   worksheet.Cell("A1").Value = "Hello World!";
   worksheet.Cell("A2").FormulaA1 = "MID(A1, 7, 5)";
   workbook.SaveAs("HelloWorld.xlsx");

.. note::
   These docs are very much a work in progress. If you'd like to contribute, click on the *Edit on GitHub* link in the right top corner.

.. toctree::
   :maxdepth: 1
   :caption: Quick Start

   installation

.. toctree::
   :maxdepth: 2
   :caption: Concepts

   concepts/types

.. toctree::
   :maxdepth: 3
   :caption: Features

   features/worksheets
   features/bulk-insert-data
   features/cell-styles

.. toctree::
   :maxdepth: 2
   :caption: API Reference

   api/index
   api/workbook
   api/worksheet
   api/cell

Contribute
----------

- Issue Tracker: https://github.com/ClosedXML/ClosedXML/issues
- Source Code: https://github.com/ClosedXML/ClosedXML
