Types
=====

A workbook in Excel is a basically a 2D array of input/output cells. Each cell can have a value with one of following types

* **Blank** - an empty value. Mostly for blanks cells or used for optional arguments of functions.
* **Logical** - a boolean value
* **Number** - an equivalent of double, but without `NaN` or `Infinity`. Number also represents dates/timespans through serial datetime.
* **Text** - A text of up to 32767 characters
* **Error** - One of the excel errors, `#DIV/0`

In the future, we will call a union of these types a **scalar value**.

Workbook also contains formulas, recipes that take a value and calculate new values. Formulas are used in

* cells formula - in that case, the output of the formula is written into the cell
* names - although names are mostly used to refer to a range of cells, they are formulas and thus can contain any formula expression (e.g. `1+2`)
* array formulas - formula changes value of 2D array of cells.

Values used during formula evaluation can have the types that are in cell, but in addition can have also following types:

* **array** - a 2D array of scalar values.
* **reference** - A reference to a range of cells, possibly non-contiguous.
