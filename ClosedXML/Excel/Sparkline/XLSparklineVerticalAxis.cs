// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSparklineVerticalAxis : IXLSparklineVerticalAxis
    {
        #region Public Properties

        public Double? ManualMax
        {
            get => _manualMax;
            set => SetManualMax(value);
        }

        public Double? ManualMin
        {
            get => _manualMin;
            set => SetManualMin(value);
        }

        public XLSparklineAxisMinMax MaxAxisType
        {
            get => _maxAxisType;
            set => SetMaxAxisType(value);
        }

        public XLSparklineAxisMinMax MinAxisType
        {
            get => _minAxisType;
            set => SetMinAxisType(value);
        }

        #endregion Public Properties

        #region Public Methods

        public IXLSparklineVerticalAxis SetManualMax(Double? manualMax)
        {
            if (manualMax != null)
                MaxAxisType = XLSparklineAxisMinMax.Custom;

            _manualMax = manualMax;
            return this;
        }

        public IXLSparklineVerticalAxis SetManualMin(Double? manualMin)
        {
            if (manualMin != null)
                MinAxisType = XLSparklineAxisMinMax.Custom;

            _manualMin = manualMin;
            return this;
        }

        public IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax maxAxisType)
        {
            if (maxAxisType != XLSparklineAxisMinMax.Custom)
                _manualMax = null;

            _maxAxisType = maxAxisType;
            return this;
        }

        public IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax minAxisType)
        {
            if (minAxisType != XLSparklineAxisMinMax.Custom)
                _manualMin = null;

            _minAxisType = minAxisType;
            return this;
        }

        #endregion Public Methods

        #region Private Fields

        private Double? _manualMax;
        private Double? _manualMin;
        private XLSparklineAxisMinMax _maxAxisType;
        private XLSparklineAxisMinMax _minAxisType;

        #endregion Private Fields

        public IXLSparklineGroup SparklineGroup { get; }

        public XLSparklineVerticalAxis(IXLSparklineGroup sparklineGroup)
        {
            SparklineGroup = sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));
        }

        public static void Copy(IXLSparklineVerticalAxis from, IXLSparklineVerticalAxis to)
        {
            to.ManualMax = from.ManualMax;
            to.ManualMin = from.ManualMin;
            to.MaxAxisType = from.MaxAxisType;
            to.MinAxisType = from.MinAxisType;
        }
    }
}
