using System;
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
        /// <param name="worksheet">The worksheet the table will live on</param>
        /// <param name="message">Message if validation fails</param>
        /// <returns>True if the proposed table name is valid in the context of the workbook</returns>
        public static bool IsValidTableNameInWorkbook(string tableName, IXLWorksheet worksheet, out string message)
        {
            message = "";

            //TODO: Technically a table name should be unique across the entire workbook not just a sheet.
            //For now I am just consolidating the table name validation logic here. If we wanted to match Excel
            //We are going to have to refactor logic around copying tables and worksheets between workbooks
            //Currently this isn't a problem because on save we call GetTableName in XL_Workbook_Save which will ensure
            //There are no conflicts in table names when we write the files, but will mean names are changed when you
            //Reopen the file.
            var existingSheetNames = worksheet.Tables.Select(t => t.Name);
            // var existingSheetNames = GetTableNamesAcrossWorkbook(worksheet.Workbook);

            //Validate common name rules, as well as check for existing conflicts
            if (!XLHelper.ValidateName("table", tableName, String.Empty, existingSheetNames, out message))
            {
                return false;
            }

            //Table names cannot contain spaces
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
            if (IsTableNameIsUniqueAcrossNamedRanges(tableName, worksheet.Workbook))
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

        // private static IList<string> GetTableNamesAcrossWorkbook(IXLWorkbook workbook)
        // {
        //     return workbook.Worksheets.SelectMany(ws => ws.Tables.Select(t => t.Name)).ToList();
        // }
    }
}
