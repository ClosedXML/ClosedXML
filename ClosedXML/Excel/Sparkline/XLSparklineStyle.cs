// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSparklineStyle : IXLSparklineStyle, IEquatable<XLSparklineStyle>
    {
        #region Public Properties

        public XLColor FirstMarkerColor { get; set; }

        public XLColor HighMarkerColor { get; set; }

        public XLColor LastMarkerColor { get; set; }

        public XLColor LowMarkerColor { get; set; }

        public XLColor MarkersColor { get; set; }

        public XLColor NegativeColor { get; set; }

        public XLColor SeriesColor { get; set; }

        #endregion Public Properties

        #region Public Methods

        public IXLSparklineStyle SetFirstMarkerColor(XLColor value)
        {
            FirstMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetHighMarkerColor(XLColor value)
        {
            HighMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetLastMarkerColor(XLColor value)
        {
            LastMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetLowMarkerColor(XLColor value)
        {
            LowMarkerColor = value;
            return this;
        }

        public IXLSparklineStyle SetMarkersColor(XLColor value)
        {
            MarkersColor = value;
            return this;
        }

        public IXLSparklineStyle SetNegativeColor(XLColor value)
        {
            NegativeColor = value;
            return this;
        }

        public IXLSparklineStyle SetSeriesColor(XLColor value)
        {
            SeriesColor = value;
            return this;
        }

        #endregion Public Methods

        public static void Copy(IXLSparklineStyle from, IXLSparklineStyle to)
        {
            to.FirstMarkerColor = from.FirstMarkerColor;
            to.HighMarkerColor = from.HighMarkerColor;
            to.LastMarkerColor = from.LastMarkerColor;
            to.LowMarkerColor = from.LowMarkerColor;
            to.MarkersColor = from.MarkersColor;
            to.NegativeColor = from.NegativeColor;
            to.SeriesColor = from.SeriesColor;
        }

        #region IEquatable implementation

        /// <summary>Returns a value that indicates whether two <see cref="T:ClosedXML.Excel.XLSparklineStyle" /> objects have different values.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
        public static bool operator !=(XLSparklineStyle left, XLSparklineStyle right)
        {
            return !Equals(left, right);
        }

        /// <summary>Returns a value that indicates whether the values of two <see cref="T:ClosedXML.Excel.XLSparklineStyle" /> objects are equal.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
        public static bool operator ==(XLSparklineStyle left, XLSparklineStyle right)
        {
            return Equals(left, right);
        }

        /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>true if the current object is equal to the <paramref name="other">other</paramref> parameter; otherwise, false.</returns>
        public bool Equals(XLSparklineStyle other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return FirstMarkerColor.Equals(other.FirstMarkerColor)
                   && HighMarkerColor.Equals(other.HighMarkerColor)
                   && LastMarkerColor.Equals(other.LastMarkerColor)
                   && LowMarkerColor.Equals(other.LowMarkerColor)
                   && MarkersColor.Equals(other.MarkersColor)
                   && NegativeColor.Equals(other.NegativeColor)
                   && SeriesColor.Equals(other.SeriesColor);
        }

        /// <summary>Determines whether the specified object is equal to the current object.</summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>true if the specified object  is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(XLSparklineStyle)) return false;
            return Equals((XLSparklineStyle)obj);
        }

        /// <summary>Serves as the default hash function.</summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = FirstMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ HighMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ LastMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ LowMarkerColor.GetHashCode();
                hashCode = (hashCode * 397) ^ MarkersColor.GetHashCode();
                hashCode = (hashCode * 397) ^ NegativeColor.GetHashCode();
                hashCode = (hashCode * 397) ^ SeriesColor.GetHashCode();
                return hashCode;
            }
        }

        #endregion IEquatable implementation
    }
}
