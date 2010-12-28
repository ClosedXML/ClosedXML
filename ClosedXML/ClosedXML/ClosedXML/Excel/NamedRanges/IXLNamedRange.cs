using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLNamedRange
    {
        /// <summary>
        /// Gets or sets the name of the range.
        /// </summary>
        /// <value>
        /// The name of the range.
        /// </value>
        String Name { get; set; }

        /// <summary>
        /// Gets the ranges associated with this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        IXLRanges Ranges { get; }

        /// <summary>
        /// Gets the single range associated with this named range.
        /// <para>An exception will be thrown if there are multiple ranges associated with this named range.</para>
        /// </summary>
        IXLRange Range { get; }

        /// <summary>
        /// Gets or sets the comment for this named range.
        /// </summary>
        /// <value>
        /// The comment for this named range.
        /// </value>
        String Comment { get; set; }

        /// <summary>
        /// Adds the specified range to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The range address to add.</param>
        IXLRanges Add(String rangeAddress);

        /// <summary>
        /// Adds the specified range to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The range to add.</param>
        IXLRanges Add(IXLRange range);

        /// <summary>
        /// Adds the specified ranges to this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The ranges to add.</param>
        IXLRanges Add(IXLRanges ranges);


        /// <summary>
        /// Deletes this named range (not the cells).
        /// </summary>
        void Delete();

        /// <summary>
        /// Clears the list of ranges associated with this named range.
        /// <para>(it does not clear the cells)</para>
        /// </summary>
        void Clear();

        /// <summary>
        /// Removes the specified range from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The range address to remove.</param>
        void Remove(String rangeAddress);

        /// <summary>
        /// Removes the specified range from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The range to remove.</param>
        void Remove(IXLRange range);

        /// <summary>
        /// Removes the specified ranges from this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        /// <param name="rangeAddress">The ranges to remove.</param>
        void Remove(IXLRanges ranges);
    }
}
