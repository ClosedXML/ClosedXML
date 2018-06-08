using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLSparklineGroups : IEnumerable<IXLSparklineGroup>
    {
        IXLSparklineGroup Add(IXLWorksheet targetWorksheet);

        IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroup, IXLWorksheet targetWorksheet);

        void RemoveAll();

        void Remove(IXLSparklineGroup sparklineGroup);

        IXLSparkline Find(IXLCell cell);

        void Remove(IXLCell cell);

        void Remove(Predicate<IXLSparklineGroup> predicate);

        void CopyTo(IXLWorksheet targetSheet);
    }
}
