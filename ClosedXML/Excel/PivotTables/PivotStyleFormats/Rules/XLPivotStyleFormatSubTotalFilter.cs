using System;

namespace ClosedXML.Excel
{
    [Flags]
    internal enum XLPivotStyleFormatSubTotalFilter
    {
        // Documented in https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_reference_topic_ID0EXAFBB.html#topic_ID0EXAFBB
        // Not all of these filters are available on the public API. We try at least to support loading and saving of filters

        None,

        /// <summary>
        /// Specifies a boolean value that indicates whether the 'average' aggregate function is included in the filter.
        /// </summary>
        AverageSubtotal = 1 << 1,

        /// <summary>
        /// Specifies a boolean value that indicates whether the 'countA' subtotal is included in the filter.
        /// </summary>
        CountASubtotal = 1 << 2,

        /// <summary>
        /// Specifies a boolean value that indicates whether the count aggregate function is included in the filter.
        /// </summary>
        CountSubtotal = 1 << 3,

        /// <summary>
        /// Specifies a boolean value that indicates whether the default subtotal is included in the filter.
        /// </summary>
        DefaultSubtotal = 1 << 4,

        /// <summary>
        /// Specifies a boolean value that indicates whether the 'maximum' aggregate function is included in the filter.
        /// </summary>
        MaxSubtotal = 1 << 5,

        /// <summary>
        /// Specifies a boolean value that indicates whether the 'minimum' aggregate function is included in the filter.
        /// </summary>
        MinSubtotal = 1 << 6,

        /// <summary>
        /// Specifies a boolean value that indicates whether the 'product' aggregate function is included in the filter.
        /// </summary>
        ProductSubtotal = 1 << 7,

        /// <summary>
        /// Specifies a boolean value that indicates whether the population standard deviation aggregate function is included in the filter.
        /// </summary>
        StandardDeviationPSubtotal = 1 << 8,

        /// <summary>
        /// Specifies a boolean value that indicates whether the standard deviation aggregate function is included in the filter.
        /// </summary>
        StandardDeviationSubtotal = 1 << 9,

        /// <summary>
        /// Specifies a boolean value that indicates whether the sum aggregate function is included in the filter.
        /// </summary>
        SumSubtotal = 1 << 10,

        /// <summary>
        /// Specifies a boolean value that indicates whether the population variance aggregate function is included in the filter.
        /// </summary>
        VariancePSubtotal = 1 << 11,

        /// <summary>
        /// Specifies a boolean value that indicates whether the variance aggregate function is included in the filter.
        /// </summary>
        VarianceSubtotal = 1 << 12,
    }
}
