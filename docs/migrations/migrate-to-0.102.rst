#############################
Migration from 0.101 to 0.102
#############################

**************
Array formulas
**************

Setting the IXLCell.FormulaA1 to a formula with braces (e.g. ``{=1+2}``)
will cause an exception during formula evaluation. The braces were previously
stripped automatically.


*************
Fallback font
*************

`DefaultGraphicEngine` used to throw an exception, when it was unable to find
requested font nor fallback font. It now uses a stripped version of a Carlito
font (Calibri metric compatible font), when no font is available.
   
For more details and consequences, see Graphic Engine page.

*****
Cells
*****

The content of the cells is now stored in sparse arrays, instead of directly in
the ``IXLCell``. That causes several changes:

The address of cells are no longer updated, when areas are deleted or inserted.
`worksheet.Cell("A4")` will always return value at *A4*, even if row 2 has been
deleted. Previously, when row was deleted, the cell then contained data from
*A3*.

Operator ==  no longer works. Use the ``Equals`` method.

.. code-block:: csharp

   var first = ws.Cell("A1");
   var second = ws.Cell("A1");
   if (first == second) {
     // no longer works
   }
   
   if (first.Equals(second))
   {
     // works
   }
