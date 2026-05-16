using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An API object for a sparkline. It doesn't contain any data, only a link to the point of
    /// the sparkline and does operations through <see cref="XLSparklineGroup"/>. It uses the
    /// cell location as an anchor and if it is no longer valid, the group should throw.
    /// </summary>
    internal class XLSparkline : IXLSparkline
    {
        private readonly XLSparklineGroup _sparklineGroup;
        private XLSheetPoint _location;

        internal XLSparkline(XLSparklineGroup sparklineGroup, XLSheetPoint location)
        {
            _sparklineGroup = sparklineGroup;
            _location = location;
        }

        public IXLCell Location
        {
            get => _sparklineGroup.GetLocation(_location);
            set => SetLocation(value);
        }

        public IXLRange? SourceData
        {
            get => _sparklineGroup.GetSparklineSourceData(_location);
            set => SetSourceData(value);
        }

        public IXLSparklineGroup SparklineGroup => _sparklineGroup;

        public IXLSparkline SetLocation(IXLCell newLocation)
        {
            if (newLocation is null)
                throw new ArgumentNullException(nameof(newLocation));

            if (newLocation.Worksheet != SparklineGroup.Worksheet)
                throw new ArgumentException("Cannot move the sparkline to a different worksheet");

            var destination = XLSheetPoint.FromCell(newLocation);
            _sparklineGroup.MoveSparkline(_location, destination);
            _location = destination;
            return this;
        }

        public IXLSparkline SetSourceData(IXLRange? sourceDataRange)
        {
            _sparklineGroup.SetSparklineSourceData(_location, sourceDataRange);
            return this;
        }
    }
}
