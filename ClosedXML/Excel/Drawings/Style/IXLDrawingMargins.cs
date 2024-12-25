using System;

namespace ClosedXML.Excel
{
    public interface IXLDrawingMargins
    {
        Boolean Automatic { get; set; }

        /// <summary>
        /// Left margin in inches.
        /// </summary>
        Double Left { get; set; }

        /// <summary>
        /// Right margin in inches.
        /// </summary>
        Double Right { get; set; }

        /// <summary>
        /// Top margin in inches.
        /// </summary>
        Double Top { get; set; }

        /// <summary>
        /// Bottom margin in inches.
        /// </summary>
        Double Bottom { get; set; }

        /// <summary>
        /// Set <see cref="Left"/>, <see cref="Top"/>, <see cref="Right"/>, <see cref="Bottom"/> margins at once.
        /// </summary>
        Double All { set; }

        IXLDrawingStyle SetAutomatic(); IXLDrawingStyle SetAutomatic(Boolean value);
        IXLDrawingStyle SetLeft(Double value);
        IXLDrawingStyle SetRight(Double value);
        IXLDrawingStyle SetTop(Double value);
        IXLDrawingStyle SetBottom(Double value);
        IXLDrawingStyle SetAll(Double value);

    }
}
