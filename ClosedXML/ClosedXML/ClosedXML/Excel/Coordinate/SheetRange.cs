using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Lightweight struct for work with sheet coordinate range
    /// </summary>
    public struct SheetRange : IEquatable<SheetRange>
    {
        #region Static
        ///<summary>
        /// Singleton instance 
        ///</summary>
// ReSharper disable RedundantDefaultFieldInitializer
        public static readonly SheetRange Empty = new SheetRange();
// ReSharper restore RedundantDefaultFieldInitializer
        #endregion
        #region Private fields
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly SheetPoint m_begin;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly SheetPoint m_end;
        #endregion
        #region Constructor
        public SheetRange(int beginRow, int beginColumn, int endRow, int endColumn)
                : this(new SheetPoint(beginRow, beginColumn), new SheetPoint(endRow, endColumn))
        {
        }
        public SheetRange(SheetPoint begin, SheetPoint end)
        {
            #region Check
            if (begin.RowNumber > end.RowNumber)
            {
                throw new ArgumentOutOfRangeException("begin", "Row part of begin coordinate must be less or equal Row part coordinate of end");
            }
            if (begin.ColumnNumber > end.ColumnNumber)
            {
                throw new ArgumentOutOfRangeException("begin", "Column part of begin coordinate must be less or equal Column part coordinate of end");
            }
            #endregion
            m_begin = begin;
            m_end = end;
        }
        #endregion
        #region Public properties
        public SheetPoint FirstAddress
        {
            [DebuggerStepThrough]
            get { return m_begin; }
        }
        public SheetPoint LastAddress
        {
            [DebuggerStepThrough]
            get { return m_end; }
        }

        public bool IsOneCell
        {
            get { return m_begin.Equals(m_end); }
        }
        public int RowCount
        {
            [DebuggerStepThrough]
            get { return m_end.RowNumber - m_begin.RowNumber + 1; }
        }
        public int ColumnCount
        {
            [DebuggerStepThrough]
            get { return m_end.ColumnNumber - m_begin.ColumnNumber + 1; }
        }
        public int Count
        {
            [DebuggerStepThrough]
            get { return RowCount*ColumnCount; }
        }
        #endregion
        #region Public methods
        public bool Equals(SheetRange other)
        {
            return other.m_begin.Equals(m_begin) && other.m_end.Equals(m_end);
        }
        public bool Intersects(SheetRange range)
        {
            return !(range.FirstAddress.ColumnNumber > LastAddress.ColumnNumber
                     || range.LastAddress.ColumnNumber < FirstAddress.ColumnNumber
                     || range.FirstAddress.RowNumber > LastAddress.RowNumber
                     || range.LastAddress.RowNumber < FirstAddress.RowNumber
                    );
        }
        
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (obj.GetType() != typeof (SheetRange))
            {
                return false;
            }
            return Equals((SheetRange) obj);
        }
        public override int GetHashCode()
        {
            unchecked
            {
                return (m_begin.GetHashCode()*397) ^ m_end.GetHashCode();
            }
        }
        internal string ToStringA1()
        {
            return IsOneCell ? m_begin.ToStringA1() : string.Format("{0}:{1}", m_begin.ToStringA1(), m_end.ToStringA1());
        }
        public override string ToString()
        {
            return ToStringA1();
        }
        #endregion
        #region Internal methods
        internal bool Equals(XLRangeAddress other)
        {
            if (ReferenceEquals(other, null))
            {
                return false;
            }
            return m_begin.Equals(other.FirstAddress) && m_end.Equals(other.LastAddress);
        }

        internal bool Intersects(XLRangeAddress range)
        {
            if (ReferenceEquals(range, null))
            {
                return false;
            }
            return !(range.FirstAddress.ColumnNumber > LastAddress.ColumnNumber
                     || range.LastAddress.ColumnNumber < FirstAddress.ColumnNumber
                     || range.FirstAddress.RowNumber > LastAddress.RowNumber
                     || range.LastAddress.RowNumber < FirstAddress.RowNumber
                    );
        }

        #endregion

        
    }
}