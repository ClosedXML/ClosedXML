using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace ClosedXML.Excel.Charts
{
    public class ChartAxis
    {
        private static UInt32 s_uniqueId = 1;

        public String Name { get; set; }
        public ChartAxisType Type { get; set; }
        public ChartAxis RelatedAxis { get; set; }
        public bool Invisible { get; set; }
        public AxisPositionValues Position { get; set; }
        public double? MinValue { get; set; }
        public double? MaxValue { get; set; }

        public bool InvertOrientation { get; set; }
        public int TickMark { get; set; }
        public bool CrossMax { get; set; }

        public CrossBetweenValues CrossBetween { get; set; }

        private UInt32 m_UniqueId { get; set; }
        internal UInt32 Id
        {
            get
            {
                if (m_UniqueId == 0)
                    m_UniqueId = s_uniqueId++;
                return m_UniqueId;
            }
        }

        public enum ChartAxisType
        {
            Undefined,
            Category,
            ValuePercent100WithAllTickmarks,
            ValueGeneric,
            FakeCategory,
            ValueDifference,
        }
    }
}
