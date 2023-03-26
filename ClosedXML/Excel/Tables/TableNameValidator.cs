using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /*
     * A string representing the name of the table. This is the name that shall be used in formula references,
     *  and displayed in the UI to the spreadsheet user. This name shall not have any spaces in it,
     * and it must be unique amongst all other displayNames and definedNames in the workbook.
     * The character lengths and restrictions are the same as for definedNames.
     * See SpreadsheetML Reference - Workbook definedNames section for details
     * The possible values for this attribute are defined by the ST_Xstring simple type (§3.18.96).
     */

    internal static class TableNameValidator
    {
        /// <summary>
        /// Validates if a suggested TableName is valid in the context of a specific workbook
        /// </summary>
        /// <param name="tableName">Proposed Table Name</param>
        /// <param name="workbook"></param>
        /// <param name="message">Message if validation fails</param>
        /// <returns>True if the proposed table name is valid in the context of the workbook</returns>
        public static bool IsValidTableNameInWorkbook(string tableName, IXLWorkbook workbook, out string message)
        {
            message = "";

            var existingSheetNames = GetTableNamesAcrossWorkbook(workbook);

            //Validate common name rules, as well as check for existing conflicts
            if (!XLHelper.ValidateName("table", tableName, String.Empty, existingSheetNames, out message))
            {
                return false;
            }

            //Perform table specific names validation
            if (tableName.Contains(" "))
            {
                message = "Table names cannot contain spaces";
                return false;
            }

            //Validate TableName is not a Cell Address
            if (XLHelper.IsValidA1Address(tableName) || XLHelper.IsValidRCAddress(tableName))
            {
                message = $"Table name cannot be a valid Cell Address '{tableName}'.";
                return false;
            }


            //A Table name must be unique across all defined names regardless of if it scoped to workbook or sheet
            if (IsTableNameIsUniqueAcrossNamedRanges(tableName, workbook))
            {
                message = $"Table name must be unique across all named ranges '{tableName}'.";
                return false;
            }

            return true;
        }

        private static bool IsTableNameIsUniqueAcrossNamedRanges(string tableName, IXLWorkbook workbook)
        {
            //Check both workbook and worksheet scoped named ranges
            return workbook.NamedRanges.Contains(tableName) ||
                   workbook.Worksheets.Any(ws => ws.NamedRanges.Contains(tableName));
        }

        /// <summary>
        /// Get all tables names in the workbook. Table names MUST be unique across the whole workbook, not just the sheet
        /// </summary>
        /// <param name="workbook">workbook context</param>
        /// <returns>String collection representing all the table names in the workbook</returns>
        private static IList<string> GetTableNamesAcrossWorkbook(IXLWorkbook workbook)
        {
            return workbook.Worksheets.SelectMany(ws => ws.Tables.Select(t => t.Name)).ToList();
        }
    }
}
