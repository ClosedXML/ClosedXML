// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    ///  Relative or absolute reference to a rectangular range
    /// </summary>
    internal class XLRangeReference : IXLReference
    {
        #region Public Properties

        public XLCellReference FirstCell { get; }

        public XLCellReference LastCell { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLRangeReference(XLCellReference firstCell, XLCellReference lastCell)
        {
            FirstCell = firstCell ?? throw new ArgumentNullException(nameof(firstCell));
            LastCell = lastCell ?? throw new ArgumentNullException(nameof(lastCell));
        }

        #endregion Public Constructors

        #region Public Methods

        public override string ToString() => ToStringR1C1();

        public string ToStringA1(IXLAddress baseAddress)
        {
            return $"{FirstCell.ToStringA1(baseAddress)}:{LastCell.ToStringA1(baseAddress)}";
        }

        public string ToStringR1C1()
        {
            return $"{FirstCell.ToStringR1C1()}:{LastCell.ToStringR1C1()}";
        }

        #endregion Public Methods
    }
}
