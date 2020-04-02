// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel.InsertData
{
    /// <summary>
    /// A universal interface for different data readers used in InsertData logic.
    /// </summary>
    internal interface IInsertDataReader
    {
        /// <summary>
        /// Get a collection of records, each as a collection of values, extracted from a source.
        /// </summary>
        IEnumerable<IEnumerable<object>> GetData();

        /// <summary>
        /// Get the number of properties to use as a table with.
        /// Actual number of may vary in different records.
        /// </summary>
        int GetPropertiesCount();

        /// <summary>
        /// Get the title of the property with the specified index.
        /// </summary>
        string GetPropertyName(int propertyIndex);

        /// <summary>
        /// Get the total number of records.
        /// </summary>
        /// <returns></returns>
        int GetRecordsCount();
    }
}
