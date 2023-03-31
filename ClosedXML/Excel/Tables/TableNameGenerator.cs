using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class TableNameGenerator
    {
        /// <summary>
        /// Generate a non conflicting table name for the workbook
        /// </summary>
        /// <param name="workbook">Workbook to generate name for</param>
        /// <param name="baseTableName">"Base Table Name</param>
        /// <returns>Name for the table</returns>
        internal static string GetNewTableName(IXLWorkbook workbook, string baseTableName = "Table")
        {
            var existingTableNames = new HashSet<String>(
                workbook.Worksheets
                    .SelectMany(ws => ws.Tables)
                    .Select(t => t.Name),
                StringComparer.OrdinalIgnoreCase);

            var i = 1;
            string tableName;
            do
            {
                tableName = baseTableName + i;
                i++;
            } while (existingTableNames.Contains(tableName));

            return tableName;
        }
    }
}
