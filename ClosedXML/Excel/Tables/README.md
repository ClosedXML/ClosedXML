# Invariants

`table/tableColumns/tableColumn[@totalsRowLabel]` attribute must match the cell value stored in the sheet at the totals row label cell. If the text in a cell doesn't match (or is missing), Excel considers the workbook to be corrupt and tries to repair it.
