Cell styles
***********

Cells can be styled, their background, border, font of the content and many other options.

Fluent vs properties
####################
The style can be set through properties or through fluent API. Both style produce same result.

.. code-block:: csharp

   // Set style through property
   ws.Cell("A1").Style.Font.FontSize = 20;
   ws.Cell("A1").Style.Font.FontName = "Arial";

   // Set style using fluent API
   ws.Cell("A1").Style
       .Font.SetFontSize(20)
       .Font.SetFontName("Arial");

Font
####

.. code-block:: csharp

   ws.Cell("A1").Style
       .Font.SetFontSize(20)
       .Font.SetFontName("Arial");

Background color
-------------------

.. code-block:: csharp

	ws.Cell("A1").Style
		.Fill.SetBackgroundColor(XLColor.Red);

-----------
Cell border
-----------
You can set a border of a cell.

.. code-block:: csharp

   // Default color is black
   ws.Cell("B2").Style
       .Border.SetTopBorder(XLBorderStyleValues.Medium)
       .Border.SetRightBorder(XLBorderStyleValues.Medium)
       .Border.SetBottomBorder(XLBorderStyleValues.Medium)
       .Border.SetLeftBorder(XLBorderStyleValues.Medium);
