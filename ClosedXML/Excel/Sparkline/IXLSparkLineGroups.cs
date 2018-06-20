using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLSparklineGroups : IEnumerable<IXLSparklineGroup>
    {
        IXLSparklineGroup Add(IXLWorksheet targetWorksheet, String name = "");

        IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroupToCopy, IXLWorksheet targetWorksheet, String name = "");

        void RemoveAll();

        void Remove(IXLSparklineGroup sparklineGroup);

        IXLSparklineGroup Find(String name);

        List<IXLSparkline> FindSparklines(IXLRangeBase rangeBase);

        IXLSparkline FindSparkline(IXLCell cell);
        
        void Remove(IXLCell cell);

        void Remove(IXLSparkline sparkline);

        void CopyTo(IXLWorksheet targetSheet, String name = "");
    }
}
