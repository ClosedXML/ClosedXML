using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroups : IXLSparklineGroups
    {
        private readonly List<IXLSparklineGroup> _sparklineGroups = new List<IXLSparklineGroup>();

        private String GetNextSparklineGroupName()
        {
            string sparklineGroupName = "SparklineGroup1";
            int i = 1;
            while(_sparklineGroups.FirstOrDefault(slg => slg.Name == sparklineGroupName) != null)
            {
                sparklineGroupName = "SparklineGroup" + i++.ToString();
            }

            return sparklineGroupName;
        }

        /// <summary>
        /// Add a new sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <param name="name">A name for this sparkline group, leave empty to assign the next available default name.</param>
        /// <returns>The new sparkline group added</returns>
        public IXLSparklineGroup Add(IXLWorksheet targetWorksheet, String name = "")
        {
            var sparklineGroup = new XLSparklineGroup(targetWorksheet, (name == "") ? GetNextSparklineGroupName() : name);
            _sparklineGroups.Add(sparklineGroup);            
            return sparklineGroup;
        }

        /// <summary>
        /// Add a copy of an existing sparkline group to the specified worksheet
        /// </summary>        
        /// <param name="sparklineGroupToCopy">The sparkline group to copy</param>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <param name="name">A name for this sparkline group, leave empty to assign the next available default name.</param>
        /// <returns>The new sparkline group added</returns>
        public IXLSparklineGroup AddCopy(IXLSparklineGroup sparklineGroupToCopy, IXLWorksheet targetWorksheet, String name = "")
        {
            var sparklineGroup = new XLSparklineGroup(sparklineGroupToCopy, targetWorksheet, (name == "") ? GetNextSparklineGroupName() : name);
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

        /// <summary>
        /// Search for the first sparkline group with the specified name
        /// </summary>
        /// <param name="name">The name to search for</param>
        /// <returns>The first sparkline group with the name or null if no sparkline groups exist with that name</returns>
        public IXLSparklineGroup Find(String name)
        {
            return this.FirstOrDefault(slg => slg.Name == name);
        }

        /// <summary>
        /// Search for the first sparkline that is in the specified cell
        /// </summary>
        /// <param name="cell">The cell to find the sparkline for</param>
        /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
        public IXLSparkline FindSparkline(IXLCell cell)
        {
            foreach (var slg in _sparklineGroups)
            {
                if (slg.Any(sl => sl.Cell == cell))
                    return slg.First(sl => sl.Cell == cell);
            }

            return null;
        }

        /// <summary>
        /// Find all sparklines located in a given range
        /// </summary>
        /// <param name="searchRange">The range to search</param>
        /// <returns>The sparkline in the cell or null if no sparklines are found</returns>
        public List<IXLSparkline> FindSparklines(IXLRangeBase searchRange)
        {
            List<IXLSparkline> sparklines = new List<IXLSparkline>();

            foreach (var slg in _sparklineGroups)
            {
                sparklines.AddRange(slg.Where(sl => sl.GetRanges().GetIntersectedRanges(searchRange.RangeAddress).Any()));
            }

            return sparklines;
        }

        /// <summary>
        /// Remove all sparklines in the specified cell
        /// </summary>
        /// <param name="cell">The cell to remove sparklines from</param>
        public void Remove(IXLCell cell)
        {
            foreach (var slg in _sparklineGroups)
            {
                slg.Remove(cell);
            }
        }

        /// <summary>
        /// Remove the sparkline group from the worksheet
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to remove</param>
        public void Remove(IXLSparklineGroup sparklineGroup)
        {
            _sparklineGroups.Remove(sparklineGroup);
        }

        /// <summary>
        /// Remove the sparkline from the worksheet
        /// </summary>
        /// <param name="sparkline">The sparkline to remove</param>
        public void Remove(IXLSparkline sparkline)
        {
            foreach (var slg in _sparklineGroups)
            {
                slg.Remove(sparkline);
            }
        }

        /// <summary>
        /// Remove all sparkline groups and their contents from the worksheet.
        /// </summary>
        public void RemoveAll()
        {
            _sparklineGroups.Clear();
        }

        /// <summary>
        /// Copy this sparkline group to a different worksheet
        /// </summary>
        /// <param name="targetSheet">The worksheet to copy the sparkline group to</param>
        /// <param name="name">A name for this sparkline group, leave empty to assign the next available default name.</param>
        public void CopyTo(IXLWorksheet targetSheet, String name = "")
        {            
            foreach (var slg in this)
            {
                slg.CopyTo(targetSheet, name);
            }
        }
    }
}
