using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public class XLRange: IXLRange
    {
        private IXLStyle defaultStyle;

        public XLRange(XLRangeParameters xlRangeParameters)
        {
            Internals = new XLRangeInternals(xlRangeParameters.FirstCellAddress, xlRangeParameters.LastCellAddress, xlRangeParameters.Worksheet);
            RowNumber = xlRangeParameters.FirstCellAddress.Row;
            ColumnNumber = xlRangeParameters.FirstCellAddress.Column;
            ColumnLetter = xlRangeParameters.FirstCellAddress.ColumnLetter;
            this.defaultStyle = new XLStyle(this, xlRangeParameters.DefaultStyle);
        }

        #region IXLRange Members

        public IXLRange Row(Int32 row)
        {
            IXLAddress firstCellAddress = new XLAddress(row, 1);
            IXLAddress lastCellAddress = new XLAddress(row, this.ColumnCount());
            return this.Range(firstCellAddress, lastCellAddress);
        }
        public IXLRange Column(Int32 column)
        {
            return this.Range(1, column, this.RowCount(), column);
        }
        public IXLRange Column(String column)
        {
            return Column(XLAddress.GetColumnNumberFromLetter(column));
        }

        public Int32 RowNumber { get; private set; }
        public Int32 ColumnNumber { get; private set; }
        public String ColumnLetter { get; private set; }

        public IXLRangeInternals Internals { get; private set; }
        
        #endregion

        #region IXLStylized Members

        public IXLStyle Style 
        {
            get
            {
                return this.defaultStyle;
            }
            set
            {
                this.Cells().ForEach(c => c.Style = value);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                foreach (var cell in this.Cells())
                {
                    yield return cell.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion
    }
}
