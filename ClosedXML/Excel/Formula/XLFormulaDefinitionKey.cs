using System;

namespace ClosedXML.Excel
{
    internal struct XLFormulaDefinitionKey : IEquatable<XLFormulaDefinitionKey>
    {
        private readonly string _formulaR1C1;

        public XLFormulaDefinitionKey(string formulaR1C1) : this()
        {
            _formulaR1C1 = formulaR1C1 ?? string.Empty;
        }

        public string FormulaR1C1 => _formulaR1C1 ?? string.Empty;

        #region Equality members

        /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>true if the current object is equal to the <paramref name="other">other</paramref> parameter; otherwise, false.</returns>
        public bool Equals(XLFormulaDefinitionKey other)
        {
            return string.Equals(_formulaR1C1, other._formulaR1C1);
        }

        /// <summary>Indicates whether this instance and a specified object are equal.</summary>
        /// <param name="obj">The object to compare with the current instance.</param>
        /// <returns>true if <paramref name="obj">obj</paramref> and this instance are the same type and represent the same value; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is XLFormulaDefinitionKey other && Equals(other);
        }

        /// <summary>Returns the hash code for this instance.</summary>
        /// <returns>A 32-bit signed integer that is the hash code for this instance.</returns>
        public override int GetHashCode()
        {
            return _formulaR1C1.GetHashCode();
        }

        /// <summary>Returns a value that indicates whether the values of two <see cref="T:ClosedXML.Excel.Formula.XLFormulaDefinitionKey" /> objects are equal.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
        public static bool operator ==(XLFormulaDefinitionKey left, XLFormulaDefinitionKey right)
        {
            return left.Equals(right);
        }

        /// <summary>Returns a value that indicates whether two <see cref="T:ClosedXML.Excel.Formula.XLFormulaDefinitionKey" /> objects have different values.</summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
        public static bool operator !=(XLFormulaDefinitionKey left, XLFormulaDefinitionKey right)
        {
            return !left.Equals(right);
        }

        #endregion
    }
}
