namespace ClosedXML.Excel
{
    /// <summary>
    /// An enum that represents types of values in pivot cache records. It represents
    /// values under <c>CT_Record</c> type.
    /// </summary>
    internal enum XLPivotCacheValueType
    {
        /// <summary>
        /// A blank value.
        /// </summary>
        Missing,

        /// <summary>
        /// Double precision number, not <c>NaN</c> or <c>infinity</c>.
        /// </summary>
        Number,

        /// <summary>
        /// Bool value.
        /// </summary>
        Boolean,

        /// <summary>
        /// <see cref="XLError"/> value.
        /// </summary>
        Error,

        /// <summary>
        /// Cache value is a string. Because references can't be converted to number (GC would not accept it),
        /// the value is an index into a table of strings in the cache.
        /// </summary>
        String,

        // TODO: Does excel allows full range of xsd:dateTime? Zulu, offsets ect?
        /// <summary>
        /// Value is a date time.
        /// </summary>
        DateTime,

        /// <summary>
        /// Value is a reference to the shared item. The index value is an
        /// index into the shared items array of the field.
        /// </summary>
        Index,
    }
}
