// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLAllowedValues { AnyValue, WholeNumber, Decimal, Date, Time, TextLength, List, Custom }

    public enum XLErrorStyle { Stop, Warning, Information }

    public enum XLOperator { EqualTo, NotEqualTo, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan, Between, NotBetween }

    public interface IXLDataValidation
    {
        XLAllowedValues AllowedValues { get; set; }

        XLDateCriteria Date { get; }

        XLDecimalCriteria Decimal { get; }

        string ErrorMessage { get; set; }

        XLErrorStyle ErrorStyle { get; set; }

        string ErrorTitle { get; set; }

        bool IgnoreBlanks { get; set; }

        bool InCellDropdown { get; set; }

        string InputMessage { get; set; }

        string InputTitle { get; set; }

        string MaxValue { get; set; }

        string MinValue { get; set; }

        XLOperator Operator { get; set; }

        /// <summary>
        /// A collection of ranges the data validation rule applies too.
        /// </summary>
        IEnumerable<IXLRange> Ranges { get; }

        bool ShowErrorMessage { get; set; }

        //void Delete();
        //void CopyFrom(IXLDataValidation dataValidation);
        bool ShowInputMessage { get; set; }

        XLTextLengthCriteria TextLength { get; }

        XLTimeCriteria Time { get; }

        string Value { get; set; }

        XLWholeNumberCriteria WholeNumber { get; }

        /// <summary>
        /// Add a range to the collection of ranges this rule applies to.
        /// If the specified range does not belong to the worksheet of the data validation
        /// rule it is transferred to the target worksheet.
        /// </summary>
        /// <param name="range">A range to add.</param>
        void AddRange(IXLRange range);

        /// <summary>
        /// Add a collection of ranges to the collection of ranges this rule applies to.
        /// Ranges that do not belong to the worksheet of the data validation
        /// rule are transferred to the target worksheet.
        /// </summary>
        /// <param name="ranges">Ranges to add.</param>
        void AddRanges(IEnumerable<IXLRange> ranges);

        void Clear();

        /// <summary>
        /// Detach data validation rule of all ranges it applies to.
        /// </summary>
        void ClearRanges();

        void Custom(string customValidation);

        bool IsDirty();

        void List(string list);

        void List(string list, bool inCellDropdown);

        void List(IXLRange range);

        void List(IXLRange range, bool inCellDropdown);

        /// <summary>
        /// Remove the specified range from the collection of range this rule applies to.
        /// </summary>
        /// <param name="range">A range to remove.</param>
        bool RemoveRange(IXLRange range);
    }
}
