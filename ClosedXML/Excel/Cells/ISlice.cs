using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An interface for methods of <see cref="Slice{TElement}"/> without specified type of an element.
    /// </summary>
    internal interface ISlice
    {
        /// <summary>
        /// Is at least one cell in the slice used?
        /// </summary>
        bool IsEmpty { get; }

        /// <summary>
        /// Get maximum used column in the slice or 0, if no column is used.
        /// </summary>
        int MaxColumn { get; }

        /// <summary>
        /// Get maximum used row in the slice or 0, if no row is used.
        /// </summary>
        int MaxRow { get; }

        /// <summary>
        /// A set of columns that have at least one used cell. Order of columns is non-deterministic.
        /// </summary>
        Dictionary<int, int>.KeyCollection UsedColumns { get; }

        /// <summary>
        /// A set of rows that have at least one used cell. Order of rows is non-deterministic.
        /// </summary>
        IEnumerable<int> UsedRows { get; }

        /// <summary>
        /// Clear all values in the range and mark them as unused.
        /// </summary>
        void Clear(Area area);

        /// <summary>
        /// Clear all values in the <paramref name="areaToDelete"/> and shift all values right of the deleted area to the deleted place.
        /// </summary>
        void DeleteAreaAndShiftLeft(Area areaToDelete);

        /// <summary>
        /// Clear all values in the <paramref name="areaToDelete"/> and shift all values below the deleted area to the deleted place.
        /// </summary>
        void DeleteAreaAndShiftUp(Area areaToDelete);

        /// <summary>
        /// Get all used points in a slice.
        /// </summary>
        /// <param name="area">Area to iterate over.</param>
        /// <param name="reverse"><c>false</c> = left to right, top to bottom. <c>true</c> = right to left, bottom to top.</param>
        IEnumerator<Point> GetEnumerator(Area area, bool reverse = false);

        /// <summary>
        /// Shift all values at the <paramref name="areaToInsert"/> and all cells below it
        /// down by <see cref="Area.Height"/> of the <paramref name="areaToInsert"/>.
        /// The insert area is cleared.
        /// </summary>
        void InsertAreaAndShiftDown(Area areaToInsert);

        /// <summary>
        /// Shift all values at the <paramref name="areaToInsert"/> and all cells right of it
        /// to the right by <see cref="Area.Width"/> of the <paramref name="areaToInsert"/>.
        /// The insert area is cleared.
        /// </summary>
        void InsertAreaAndShiftRight(Area areaToInsert);

        /// <summary>
        /// Does slice contains a non-default value at specified point?
        /// </summary>
        bool IsUsed(Point address);

        /// <summary>
        /// Swap content of two points.
        /// </summary>
        void Swap(Point sp1, Point sp2);
    }
}
