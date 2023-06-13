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
