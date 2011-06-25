using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Lightweight struct for work with sheet coordinates
    /// </summary>
    public struct SheetPoint : IEquatable<SheetPoint>
    {
        #region Static
        ///<summary>
        /// Singleton instance 
        ///</summary>
// ReSharper disable RedundantDefaultFieldInitializer
        public static readonly SheetPoint Empty = new SheetPoint();
// ReSharper restore RedundantDefaultFieldInitializer
        #endregion
        #region Private fields
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int m_row;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int m_column;
        #endregion
        #region Constructor
        public SheetPoint(int row, int column)
        {
            #region Check
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException("row", "Must be more than 0");
            }
            if (column < 0)
            {
                throw new ArgumentOutOfRangeException("column", "Must be more than 0");
            }
            #endregion
            m_row = row;
            m_column = column;
        }
        #endregion
        #region Public properties
        public int RowNumber
        {
            [DebuggerStepThrough]
            get { return m_row; }
        }
        public int ColumnNumber
        {
            [DebuggerStepThrough]
            get { return m_column; }
        }
        #endregion
        #region Public methods
        public bool Equals(SheetPoint other)
        {
            return other.m_row == m_row && other.m_column == m_column;
        }
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (obj.GetType() != typeof (SheetPoint))
            {
                return false;
            }
            return Equals((SheetPoint) obj);
        }
        public override int GetHashCode()
        {
            unchecked
            {
                return (m_row*397) ^ m_column;
            }
        }
        public override string ToString()
        {
            return ToStringA1();
        }
        public string ToStringA1()
        {
            return string.Format("{0}{1}", ExcelHelper.GetColumnLetterFromNumber(m_column), m_row);
        }
        #endregion
        #region Internal methods
        internal bool Equals(XLAddress other)
        {
            if (ReferenceEquals(other,null))
            {
                return false;
            }
            return m_row == other.RowNumber && m_column == other.ColumnNumber;
        }
        #endregion
    }

}
