using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// A writer for table definition part.
    /// </summary>
    internal class TablePartWriter
    {
        internal static void SynchronizeTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            // Remove table definition parts that are not a part of workbook
            foreach (var tableDefinitionPart in worksheetPart.GetPartsOfType<TableDefinitionPart>().ToList())
            {
                var partId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                var xlWorkbookContainsTable = tables.Cast<XLTable>().Any(t => t.RelId == partId);
                if (!xlWorkbookContainsTable)
                {
                    worksheetPart.DeletePart(tableDefinitionPart);
                }
            }

            foreach (var xlTable in tables.Cast<XLTable>())
            {
                if (String.IsNullOrEmpty(xlTable.RelId))
                {
                    xlTable.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    worksheetPart.AddNewPart<TableDefinitionPart>(xlTable.RelId);
                }
            }
        }

        internal static void GenerateTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            foreach (var xlTable in tables.Cast<XLTable>())
            {
                var relId = xlTable.RelId;
                var tableDefinitionPart = (TableDefinitionPart)worksheetPart.GetPartById(relId);
                GenerateTableDefinitionPartContent(tableDefinitionPart, xlTable, context);
            }
        }

        private static void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable, SaveContext context)
        {
            context.TableId++;
            var reference = xlTable.RangeAddress.FirstAddress + ":" + xlTable.RangeAddress.LastAddress;
            var tableName = GetTableName(xlTable.Name, context);
            var table = new Table
            {
                Id = context.TableId,
                Name = tableName,
                DisplayName = tableName,
                Reference = reference
            };

            if (!xlTable.ShowHeaderRow)
                table.HeaderRowCount = 0;

            if (xlTable.ShowTotalsRow)
                table.TotalsRowCount = 1;
            else
                table.TotalsRowShown = false;

            var tableColumns = new TableColumns { Count = (UInt32)xlTable.ColumnCount() };

            UInt32 columnId = 0;
            foreach (var xlField in xlTable.Fields)
            {
                columnId++;
                var fieldName = xlField.Name;
                var tableColumn = new TableColumn
                {
                    Id = columnId,
                    Name = fieldName.Replace("_x000a_", "_x005f_x000a_").Replace(Environment.NewLine, "_x000a_")
                };

                // OI-29500: the headerRowDxfId attribute shall be absent if the enclosing table element's headerRow
                // attribute value is greater than or equal to 1. The header dxf stores the format while the table
                // doesn't have a header, but user might want to add it back.
                if (xlField.HeaderFormatValue is { } headerDfx)
                    tableColumn.HeaderRowDifferentialFormattingId = context.GetDxfId(headerDfx);

                if (xlField.DataFormatValue is { } dataDfx)
                    tableColumn.DataFormatId = context.GetDxfId(dataDfx);

                if (xlField.TotalFormatValue is { } totalsDfx)
                    tableColumn.TotalsRowDifferentialFormattingId = context.GetDxfId(totalsDfx);

                // TODO Styles: Deal with this behavior, either remove or make it work.
                /*
                // https://github.com/ClosedXML/ClosedXML/issues/513
                if (xlField.IsConsistentStyle())
                {
                    var firstDataCell = (XLCell)xlField.Column.Cells()
                        .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                        .First();
                    var format = firstDataCell.FormatValue;
                    if (format is not null && context.TryGetDxfId(format, out var dxfId))
                        tableColumn.DataFormatId = dxfId;
                }
                else
                    tableColumn.DataFormatId = null;
                */

                if (xlField.IsConsistentFormula())
                {
                    string formula = xlField.Column.Cells()
                        .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                        .First()
                        .FormulaA1;

                    while (formula.StartsWith("=") && formula.Length > 1)
                        formula = formula.Substring(1);

                    if (!String.IsNullOrWhiteSpace(formula))
                    {
                        tableColumn.CalculatedColumnFormula = new CalculatedColumnFormula
                        {
                            Text = formula
                        };
                    }
                }
                else
                    tableColumn.CalculatedColumnFormula = null;

                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn.TotalsRowFunction = xlField.TotalsRowFunction.ToOpenXml();

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
                    }

                    if (!String.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn.TotalsRowLabel = xlField.TotalsRowLabel;
                }
                tableColumns.AppendChild(tableColumn);
            }

            var tableStyleInfo1 = new TableStyleInfo
            {
                ShowFirstColumn = xlTable.EmphasizeFirstColumn,
                ShowLastColumn = xlTable.EmphasizeLastColumn,
                ShowRowStripes = xlTable.ShowRowStripes,
                ShowColumnStripes = xlTable.ShowColumnStripes
            };

            if (xlTable.Theme != XLTableTheme.None)
                tableStyleInfo1.Name = xlTable.Theme.Name;

            if (xlTable.ShowAutoFilter)
            {
                var autoFilter1 = new AutoFilter();
                if (xlTable.ShowTotalsRow)
                {
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                        xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                        xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                }
                else
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);

                WorksheetPartWriter.PopulateAutoFilter(xlTable.AutoFilter, autoFilter1, context);

                table.AppendChild(autoFilter1);
            }

            table.AppendChild(tableColumns);
            table.AppendChild(tableStyleInfo1);

            tableDefinitionPart.Table = table;
        }

        private static string GetTableName(String originalTableName, SaveContext context)
        {
            var tableName = originalTableName.RemoveSpecialCharacters();
            var name = tableName;
            if (context.TableNames.Contains(name))
            {
                var i = 1;
                name = tableName + i.ToInvariantString();
                while (context.TableNames.Contains(name))
                {
                    i++;
                    name = tableName + i.ToInvariantString();
                }
            }

            context.TableNames.Add(name);
            return name;
        }
    }
}
