#############################
Migration from 0.101 to 0.102
#############################

**************
Array formulas
**************

Setting the IXLCell.FormulaA1 to a formula with braces (e.g. ``{=1+2}``)
will cause an exception during formula evaluation. The braces were previously
stripped automatically.
