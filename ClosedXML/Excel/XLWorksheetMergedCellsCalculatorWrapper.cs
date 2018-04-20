using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A wrapper class which calculates the merged cells of a collection of worksheets
    /// and resets them in the <see cref="Dispose"/> method
    /// </summary>
    internal class XLWorksheetMergedCellsCalculatorWrapper : IDisposable
    {
        private readonly List<XLWorksheet> _worksheets;

        public XLWorksheetMergedCellsCalculatorWrapper(IEnumerable<XLColumn> columns)
        {
            _worksheets = columns.Select(c => c.Worksheet).Distinct().ToList();
            _worksheets.ForEach(w => w.Internals.MergedRanges.CalculateMergedCells());
        }

        public void Dispose()
        {
            _worksheets.ForEach(w => w.Internals.MergedRanges.ResetMergedCells());
        }
    }
}
