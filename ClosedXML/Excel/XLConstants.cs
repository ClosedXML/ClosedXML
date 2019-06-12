// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    //Use the class to store magic strings or variables.
    public static class XLConstants
    {
        public const int NumberOfBuiltInStyles = 164; // But they are stored as 0-based (0 - 163)

        #region Pivot Table constants

        public const byte PivotTableCreatedVersion = 5;
        public const byte PivotTableRefreshedVersion = 5;
        public const string PivotTableValuesSentinalLabel = "{{Values}}";

        #endregion Pivot Table constants

        internal static class Comment
        {
            internal const string AlternateShapeTypeId = "_xssf_cell_comment";
            internal const string ShapeTypeId = "_x0000_t202";
        }
    }
}
