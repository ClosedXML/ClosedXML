#nullable disable

using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Parser;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPrintAreas : IXLPrintAreas, IWorkbookListener
    {
        public string PrintArea { get; private set; }

        private XLWorksheet worksheet;

        public XLPrintAreas(XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet worksheet)
        {
            PrintArea = defaultPrintAreas.PrintArea;
            this.worksheet = worksheet;
        }

        public void Clear()
        {
            PrintArea = null;
        }

        public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            AddRange(worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
        }

        public void Add(string rangeAddress)
        {
            AddRange(worksheet.Range(rangeAddress));
        }

        public void Add(string firstCellAddress, string lastCellAddress)
        {
            AddRange(worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            AddRange(worksheet.Range(firstCellAddress, lastCellAddress));
        }

        private void AddRange(XLRange range)
        {
            AddFormula(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
        }

        public void AddFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(PrintArea))
                PrintArea = formula;
            else
                PrintArea += "," + formula;
        }

        public void OnSheetRenamed(string oldSheetName, string newSheetName)
        {
            if (!string.IsNullOrWhiteSpace(PrintArea))
            {
                PrintArea = FormulaConverter.ModifyA1(PrintArea, 1, 1, new RenameRefModVisitor
                {
                    Sheets = new Dictionary<string, string> { { oldSheetName, newSheetName } }
                });
            }
        }
    }
}
