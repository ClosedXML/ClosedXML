// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLMeasureUnit
    {
        Inches,
        Millimetres,
        EnglishMetricUnits,
        Pixels
    }

    public interface IXLMeasure : IEqualityComparer<IXLMeasure>, IEquatable<IXLMeasure>
    {
        XLMeasureUnit Unit { get; }
        Double Value { get; }

        IXLMeasure ConvertTo(XLMeasureUnit unit);
    }

    // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
    // http://archive.oreilly.com/pub/post/what_is_an_emu.html
    // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
    public struct XLMeasure : IXLMeasure, IEquatable<XLMeasure>
    {
        internal readonly static IXLMeasure Zero = new XLMeasure(0, XLMeasureUnit.EnglishMetricUnits);

        internal readonly Int64 EnglishMetricUnits;
        private const Int64 emuPerInch = 914400L;
        private const Double mmPerInch = 25.4;
        private Double? value;

        internal XLMeasure(Double value, XLMeasureUnit unit)
            : this(GetEMU(value, unit))
        {
            this.Unit = unit;
        }

        private XLMeasure(Int64 emu)
        {
            this.EnglishMetricUnits = emu;
            this.Unit = XLMeasureUnit.EnglishMetricUnits;
            this.value = null;
        }

        #region Overrides

        public override string ToString()
        {
            return $"{Value} {Unit.ToString().ToLowerInvariant()}";
        }

        #endregion Overrides

        public XLMeasureUnit Unit { get; }

        public Double Value
        {
            get
            {
                if (!value.HasValue)
                {
                    value = Convert(EnglishMetricUnits, XLMeasureUnit.EnglishMetricUnits, this.Unit);
                    switch (this.Unit)
                    {
                        case XLMeasureUnit.EnglishMetricUnits:
                        case XLMeasureUnit.Pixels:
                            value = System.Convert.ToInt64(value.Value);
                            break;
                    }
                }

                return value.Value;
            }
        }

        public static XLMeasure Create(Double value, XLMeasureUnit unit)
        {
            return new XLMeasure(value, unit);
        }

        public IXLMeasure ConvertTo(XLMeasureUnit unit)
        {
            return new XLMeasure(Convert(this.EnglishMetricUnits, XLMeasureUnit.EnglishMetricUnits, unit), unit);
        }

        private static Double Convert(Double fromValue, XLMeasureUnit fromUnit, XLMeasureUnit toUnit)
        {
            if (fromUnit == toUnit)
                return fromValue;

            Double emu;
            switch (fromUnit)
            {
                case XLMeasureUnit.EnglishMetricUnits:
                    emu = fromValue;
                    break;

                case XLMeasureUnit.Inches:
                    emu = fromValue * emuPerInch;
                    break;

                case XLMeasureUnit.Millimetres:
                    emu = fromValue / mmPerInch * emuPerInch;
                    break;

                case XLMeasureUnit.Pixels:
                    emu = fromValue / XLHelper.DpiX * emuPerInch;
                    break;

                default:
                    throw new NotImplementedException();
            }

            switch (toUnit)
            {
                case XLMeasureUnit.EnglishMetricUnits:
                    return emu;

                case XLMeasureUnit.Inches:
                    return emu / emuPerInch;

                case XLMeasureUnit.Millimetres:
                    return emu / emuPerInch * mmPerInch;

                case XLMeasureUnit.Pixels:
                    return emu / emuPerInch * XLHelper.DpiX;

                default:
                    throw new NotImplementedException();
            }
        }

        private static Int64 GetEMU(Double fromValue, XLMeasureUnit fromUnit)
        {
            return System.Convert.ToInt64(Convert(fromValue, fromUnit, XLMeasureUnit.EnglishMetricUnits));
        }

        #region Operator overloads

        public static Boolean operator !=(XLMeasure left, XLMeasure right)
        {
            return !(left == right);
        }

        public static Boolean operator ==(XLMeasure left, XLMeasure right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }
            return !ReferenceEquals(left, null) && left.Equals(right);
        }

        #endregion Operator overloads

        #region IEqualityComparer<IXLMeasure> members

        public bool Equals(IXLMeasure x, IXLMeasure y)
        {
            return x == y;
        }

        public new bool Equals(object x, object y)
        {
            return x == y;
        }

        #endregion IEqualityComparer<IXLMeasure> members

        #region IEquitable<XLMeasure> members

        public bool Equals(IXLMeasure other)
        {
            if (other == null)
                return false;

            return this.EnglishMetricUnits == other.ConvertTo(XLMeasureUnit.EnglishMetricUnits).Value;
        }

        public bool Equals(XLMeasure other)
        {
            return this.EnglishMetricUnits == other.EnglishMetricUnits;
        }

        public override bool Equals(Object obj)
        {
            return Equals(obj as IXLMeasure);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public int GetHashCode(IXLMeasure obj)
        {
            var hashCode = 2122234362;
            hashCode = hashCode * -1521134295 + EnglishMetricUnits.GetHashCode();
            return hashCode;
        }

        #endregion IEquitable<XLMeasure> members
    }
}
