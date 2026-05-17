// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroups : IXLSparklineGroups, IEnumerable<XLSparklineGroup>
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLSparklineGroup> _sparklineGroups = new();

        public XLSparklineGroups(XLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        public IXLWorksheet Worksheet => _worksheet;

        /// <summary>
        /// Add empty sparkline group.
        /// </summary>
        internal XLSparklineGroup Add()
        {
            var emptyGroup = new XLSparklineGroup(Worksheet);
            Add(emptyGroup);
            return emptyGroup;
        }

        /// <summary>
        /// Add the sparkline group to the collection.
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to add to the collection</param>
        /// <returns>The same sparkline group</returns>
        public IXLSparklineGroup Add(IXLSparklineGroup sparklineGroup)
        {
            if (sparklineGroup.Worksheet != Worksheet)
                throw new ArgumentException("The specified sparkline group belongs to the different worksheet");

            _sparklineGroups.Add((XLSparklineGroup)sparklineGroup);
            return sparklineGroup;
        }

        public IXLSparklineGroup Add(string locationAddress, string sourceDataAddress)
        {
            if (locationAddress is null)
                throw new ArgumentNullException(nameof(locationAddress));

            if (sourceDataAddress is null)
                throw new ArgumentNullException(nameof(sourceDataAddress));

            return Add(new XLSparklineGroup(Worksheet, locationAddress, sourceDataAddress));
        }

        public IXLSparklineGroup Add(IXLCell location, IXLRange sourceData)
        {
            if (location is null)
                throw new ArgumentNullException(nameof(location));

            return Add(new XLSparklineGroup(location, sourceData));
        }

        public IXLSparklineGroup Add(IXLRange locationRange, IXLRange sourceDataRange)
        {
            if (locationRange is null)
                throw new ArgumentNullException(nameof(locationRange));

            if (sourceDataRange is null)
                throw new ArgumentNullException(nameof(sourceDataRange));

            return Add(new XLSparklineGroup(locationRange, sourceDataRange));
        }

        /// <summary>
        /// Add a copy of an existing sparkline group to the specified worksheet
        /// </summary>
        /// <param name="sparklineGroupToCopy">The sparkline group to copy</param>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <returns>The new sparkline group added</returns>
        public IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroupToCopy, IXLWorksheet targetWorksheet)
        {
            var sparklineGroup = new XLSparklineGroup(targetWorksheet, sparklineGroupToCopy);
            _sparklineGroups.Add(sparklineGroup);
            return sparklineGroup;
        }

        /// <summary>
        /// Copy this sparkline group to a different worksheet
        /// </summary>
        /// <param name="targetSheet">The worksheet to copy the sparkline group to</param>
        public void CopyTo(IXLWorksheet targetSheet)
        {
            foreach (var slg in this)
            {
                slg.CopyTo((XLWorksheet)targetSheet);
            }
        }

        /// <summary>
        /// Search for the first sparkline that is in the specified cell
        /// </summary>
        /// <param name="cell">The cell to find the sparkline for</param>
        /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
        public IXLSparkline? GetSparkline(IXLCell cell)
        {
            if (cell.Worksheet != _worksheet)
                return null;

            var location = XLSheetPoint.FromCell(cell);
            foreach (var sparklineGroup in _sparklineGroups)
            {
                if (sparklineGroup.TryGetSparkline(location, out var sparkline))
                    return sparkline;
            }

            return null;
        }

        /// <summary>
        /// Find all sparklines located in a given range
        /// </summary>
        /// <param name="searchRange">The range to search</param>
        /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
        public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange)
        {
            return _sparklineGroups
                .SelectMany(g => g.GetSparklines(searchRange));
        }

        public IEnumerator<XLSparklineGroup> GetEnumerator()
        {
            return _sparklineGroups.GetEnumerator();
        }

        IEnumerator<IXLSparklineGroup> IEnumerable<IXLSparklineGroup>.GetEnumerator()
        {
            return GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Remove all sparklines in the specified cell
        /// </summary>
        /// <param name="cell">The cell to remove sparklines from</param>
        public void Remove(IXLCell cell)
        {
            _sparklineGroups
                .AsParallel()
                .ForEach(g => g.Remove(cell));
        }

        public void Remove(IXLRangeBase range)
        {
            var sparklinesToRemove = _sparklineGroups
                .SelectMany(g => g)
                .Where(sparkline => range.Contains(sparkline.Location))
                .ToList();

            sparklinesToRemove.ForEach(Remove);
        }

        /// <summary>
        /// Remove the sparkline group from the worksheet
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to remove</param>
        public void Remove(IXLSparklineGroup sparklineGroup)
        {
            _sparklineGroups.Remove((XLSparklineGroup)sparklineGroup);
        }

        internal void Remove(XLSheetPoint location)
        {
            _sparklineGroups.ForEach(g => g.Remove(location));
        }

        /// <summary>
        /// Remove the sparkline from the worksheet
        /// </summary>
        /// <param name="sparkline">The sparkline to remove</param>
        private void Remove(IXLSparkline sparkline)
        {
            sparkline.SparklineGroup.Remove(sparkline);
        }

        /// <summary>
        /// Remove all sparkline groups and their contents from the worksheet.
        /// </summary>
        public void RemoveAll()
        {
            _sparklineGroups.Clear();
        }
    }
}
