using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLRangeParameters
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public
        public IXLAddress FirstCellAddress { get; set; }
        public IXLAddress LastCellAddress { get; set; }
        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; set; }
        public List<String> MergedCells { get; set; }

        // Private

        // Override


        #endregion

        #region Constructors

        // Public
        public XLRangeParameters()
        {

        }


        // Private


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public

        // Private

        // Override


        #endregion
    }
}
