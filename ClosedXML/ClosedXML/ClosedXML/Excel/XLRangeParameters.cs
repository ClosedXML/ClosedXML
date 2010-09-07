using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

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
        public Dictionary<IXLAddress, IXLCell> CellsCollection { get; set; }
        public List<String> MergedCells { get; set; }
        public IXLStyle DefaultStyle { get; set; }
        public IXLRange PrintArea { get; set; }

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
