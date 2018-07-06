// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLSparklineGroups : IEnumerable<IXLSparklineGroup>
    {
        #region Public Properties

        IXLWorksheet Worksheet { get; }

        #endregion Public Properties

        #region Public Methods

        IXLSparklineGroup Add(IXLSparklineGroup sparklineGroup);
        IXLSparklineGroup Add(string locationAddress, string sourceDataAddress);
        IXLSparklineGroup Add(IXLCell location, IXLRange sourceData);
        IXLSparklineGroup Add(IXLRange locationRange, IXLRange sourceDataRange);

        void CopyTo(IXLWorksheet targetSheet);

        IXLSparkline GetSparkline(IXLCell cell);
        IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase rangeBase);

        void Remove(IXLCell cell);
        void Remove(IXLRangeBase range);
        void Remove(IXLSparklineGroup sparklineGroup);
        void RemoveAll();

        #endregion Public Methods
    }
}
