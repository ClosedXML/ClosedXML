namespace ClosedXML.Excel
{
    /// <summary>
    /// Reference point of date/number formats available.
    /// See more at: https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
    /// </summary>
    public static class XLPredefinedFormat
    {
        public enum Number
        {
            /// <summary>
            /// General
            /// </summary>
            General = 0,

            /// <summary>
            /// 0
            /// </summary>
            Integer = 1,

            /// <summary>
            /// 0.00
            /// </summary>
            Precision2 = 2,

            /// <summary>
            /// #,##0
            /// </summary>
            IntegerWithSeparator = 3,

            /// <summary>
            /// #,##0.00
            /// </summary>
            Precision2WithSeparator = 4,

            /// <summary>
            /// 0%
            /// </summary>
            PercentInteger = 9,
        
            /// <summary>
            /// 0.00%
            /// </summary>
            PercentPrecision2 = 10,
        
            /// <summary>
            /// 0.00E+00
            /// </summary>
            ScientificPrecision2 = 11,
        
            /// <summary>
            /// # ?/?
            /// </summary>
            FractionPrecision1 = 12,
        
            /// <summary>
            /// # ??/??
            /// </summary>
            FractionPrecision2 = 13,

            /// <summary>
            /// #,##0 ,(#,##0)
            /// </summary>
            IntegerWithSeparatorAndParens = 37,

            /// <summary>
            /// #,##0 ,[Red](#,##0)
            /// </summary>
            IntegerWithSeparatorAndParensRed = 38,
        
            /// <summary>
            /// #,##0.00,(#,##0.00)
            /// </summary>
            Precision2WithSeparatorAndParens = 39,

            /// <summary>
            /// #,##0.00,[Red](#,##0.00)
            /// </summary>
            Precision2WithSeparatorAndParensRed = 40,

            /// <summary>
            /// ##0.0E+0
            /// </summary>
            ScientificUpToHundredsAndPrecision1 = 48,

            /// <summary>
            /// @
            /// </summary>
            Text = 49
        }

        public enum DateTime
        {
            /// <summary>
            /// General
            /// </summary>
            General = 0,

            /// <summary>
            /// d/m/yyyy
            /// </summary>
            DayMonthYear4WithSlashes = 14,
        
            /// <summary>
            /// d-mmm-yy
            /// </summary>
            DayMonthAbbrYear2WithDashes = 15,
        
            /// <summary>
            /// d-mmm
            /// </summary>
            DayMonthAbbrWithDash = 16,
        
            /// <summary>
            /// mmm-yy
            /// </summary>
            MonthAbbrYear2WithDash = 17,
        
            /// <summary>
            /// h:mm tt
            /// </summary>
            Hour12MinutesAmPm = 18,
        
            /// <summary>
            /// h:mm:ss tt
            /// </summary>
            Hour12MinutesSecondsAmPm = 19,
        
            /// <summary>
            /// H:mm
            /// </summary>
            Hour24Minutes = 20,
        
            /// <summary>
            /// H:mm:ss
            /// </summary>
            Hour24MinutesSeconds = 21,
        
            /// <summary>
            /// m/d/yyyy H:mm
            /// </summary>
            MonthDayYear4WithDashesHour24Minutes = 22,

            /// <summary>
            /// mm:ss
            /// </summary>
            MinutesSeconds = 45,

            /// <summary>
            /// [h]:mm:ss
            /// </summary>
            Hour12MinutesSeconds = 46,
        
            /// <summary>
            /// mmss.0
            /// </summary>
            MinutesSecondsMillis1 = 47,
        
            /// <summary>
            /// @
            /// </summary>
            Text = 49
        }
    }
}
