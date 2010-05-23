using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Allows you to add, access, and remove worksheets from the workbook.
    /// </summary>
    public class XLWorksheets: IEnumerable<XLRange>
    {
        #region Constants

        private const UInt32 MaxNumberOfRows = 1048576;
        private const UInt32 MaxNumberOfColumns = 16384;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="XLWorksheets"/> class.
        /// </summary>
        /// <param name="workbook">The workbook which will contain the worksheets.</param>
        public XLWorksheets(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        #endregion

        #region Properties

        private Dictionary<String, XLRange> worksheets = new Dictionary<String, XLRange>();

        private XLWorkbook workbook;

        /// <summary>
        /// Gets the number of worksheets in the workbook.
        /// </summary>
        public UInt32 Count
        {
            get
            {
                return (UInt32)worksheets.Count;
            }
        }

        /// <summary>
        /// Gets the worksheet (as an XLRange) with the specified sheet name.
        /// </summary>
        /// <value></value>
        public XLRange this[String sheetName]
        {
            get
            {
                return worksheets[sheetName];
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Adds a new worksheet to the workbook.
        /// </summary>
        /// <param name="name">The name of the worksheet to be added.</param>
        public XLRange Add(String name)
        {
            var firstCellAddress = new XLAddress(1, 1);
            var lastCellAddress = new XLAddress(MaxNumberOfRows, MaxNumberOfColumns);
            XLRange worksheet = new XLRange(firstCellAddress, lastCellAddress, null, null, name, workbook);
            worksheets.Add(name, worksheet);
            return worksheet;
        }

        /// <summary>
        /// Deletes the specified worksheet from the workbook.
        /// </summary>
        /// <param name="name">The name of the worksheet to be deleted.</param>
        public void Delete(String name)
        {
            worksheets.Remove(name);
        }

        #endregion


        #region IEnumerable<XLRange> Members

        public IEnumerator<XLRange> GetEnumerator()
        {
            return worksheets.Values.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion
    }
}
