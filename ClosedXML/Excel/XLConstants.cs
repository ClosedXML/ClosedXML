// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    //Use the class to store magic strings or variables.
    public static class XLConstants
    {
        internal const int NumberOfBuiltInStyles = 164; // But they are stored as 0-based (0 - 163)
        public static readonly string NewLine = "\r\n";

        #region Pivot Table constants

        public static class PivotTable
        {
            internal const byte CreatedVersion = 5;
            internal const byte RefreshedVersion = 5;

            //TODO: Needs to be refactored to be more user-friendly.
            public const string ValuesSentinalLabel = "{{Values}}";
        }

        #endregion Pivot Table constants

        internal static class Comment
        {
            internal const string AlternateShapeTypeId = "_xssf_cell_comment";
            internal const string ShapeTypeId = "_x0000_t202";
        }
    }
}