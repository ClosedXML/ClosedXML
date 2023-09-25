namespace ClosedXML.Excel
{
    /// <summary>
    /// An interface for components reacting on changes in a worksheet.
    /// </summary>
    internal interface ISheetListener
    {
        /// <summary>
        /// A handler called after the area was put into the sheet and cells shifted down.
        /// </summary>
        /// <param name="sheet">Sheet where change happened.</param>
        /// <param name="area">Area that has been inserted. The original cells were shifted down.</param>
        void OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange area);

        /// <summary>
        /// A handler called after the area was put into the sheet and cells shifted right.
        /// </summary>
        /// <param name="sheet">Sheet where change happened.</param>
        /// <param name="area">Area that has been inserted. The original cells were shifted right.</param>
        void OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange area);

        /// <summary>
        /// A handler called after the area was deleted from the sheet and cells shifted left.
        /// </summary>
        /// <param name="sheet">Sheet where change happened.</param>
        /// <param name="deletedRange">Range that has been deleted and cells to the right were shifted left.</param>
        void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedRange);

        /// <summary>
        /// A handler called after the area was deleted from the sheet and cells shifted up.
        /// </summary>
        /// <param name="sheet">Sheet where change happened.</param>
        /// <param name="deletedRange">Range that has been deleted and cells below were shifted up.</param>
        void OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedRange);
    }
}
