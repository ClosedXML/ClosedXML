using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroups : IXLSparklineGroups
    {
        private readonly List<IXLSparklineGroup> _sparklineGroups = new List<IXLSparklineGroup>();

        public IXLSparklineGroup Add(IXLWorksheet targetWorksheet)
        {
            var sparklineGroup = new XLSparklineGroup(targetWorksheet);
            _sparklineGroups.Add(sparklineGroup);            
            return sparklineGroup;
        }

        public IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroupToCopy, IXLWorksheet targetWorksheet)
        {
            var sparklineGroup = new XLSparklineGroup(sparklineGroupToCopy, targetWorksheet);
            _sparklineGroups.Add(sparklineGroup);
            return sparklineGroup;
        }

        public IEnumerator<IXLSparklineGroup> GetEnumerator()
        {
            return _sparklineGroups.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLSparkline Find(IXLCell cell)
        {
            foreach (var slg in _sparklineGroups)
            {
                if (slg.Any(sl => sl.Cell == cell))
                    return slg.First(sl => sl.Cell == cell);
            }

            return null;
        }

        public void Remove(IXLCell cell)
        {
            foreach (var slg in _sparklineGroups)
            {
                slg.Remove(cell);
            }
        }

        public void Remove(IXLSparklineGroup sparklineGroup)
        {
            _sparklineGroups.Remove(sparklineGroup);
        }

        public void Remove(Predicate<IXLSparklineGroup> predicate)
        {
            _sparklineGroups.RemoveAll(predicate);
        }

        public void RemoveAll()
        {
            _sparklineGroups.Clear();
        }

        public void CopyTo(IXLWorksheet targetSheet)
        {
            foreach (var slg in this)
            {
                slg.CopyTo(targetSheet);
            }
        }
    }
}
