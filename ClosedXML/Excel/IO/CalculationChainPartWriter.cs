#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class CalculationChainPartWriter
    {
        internal static void GenerateContent(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
        {
            if (workbookPart.CalculationChainPart == null)
                workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

            if (workbookPart.CalculationChainPart.CalculationChain == null)
                workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();

            var calculationChain = workbookPart.CalculationChainPart.CalculationChain;
            calculationChain.RemoveAllChildren<CalculationCell>();

            foreach (var worksheet in workbook.WorksheetsInternal)
            {
                foreach (var c in worksheet.Internals.CellsCollection.GetCells().Where(c => c.HasFormula))
                {
                    if (c.Formula.Type == FormulaType.DataTable)
                    {
                        // Do nothing, Excel doesn't generate calc chain for data table
                    }
                    else if (c.HasArrayFormula)
                    {
                        if (c.FormulaReference == null)
                            c.FormulaReference = c.AsRange().RangeAddress;

                        if (c.FormulaReference.FirstAddress.Equals(c.Address))
                        {
                            var cc = new CalculationCell
                            {
                                CellReference = c.Address.ToString(),
                                SheetId = (Int32)worksheet.SheetId
                            };

                            cc.Array = true;
                            calculationChain.AppendChild(cc);

                            foreach (var childCell in worksheet.Range(c.FormulaReference).Cells())
                            {
                                calculationChain.AppendChild(new CalculationCell
                                {
                                    CellReference = childCell.Address.ToString(),
                                    SheetId = (Int32)worksheet.SheetId,
                                });
                            }
                        }
                    }
                    else
                    {
                        calculationChain.AppendChild(new CalculationCell
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (Int32)worksheet.SheetId
                        });
                    }
                }
            }

            if (!calculationChain.Any())
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }
    }
}
