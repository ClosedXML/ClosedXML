namespace ClosedXML.Excel
{
    /// <summary>
    /// A very lightweight interface for entities that have an address as
    /// a rectangular range.
    /// </summary>
    public interface IXLAddressable
    {
        /// <summary>
        ///   Gets an object with the boundaries of this range.
        /// </summary>

        IXLRangeAddress RangeAddress { get; }
    }
}
