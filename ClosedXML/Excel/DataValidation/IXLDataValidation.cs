using System;

namespace ClosedXML.Excel
{
    public enum XLErrorStyle { Stop, Warning, Information }
    public enum XLAllowedValues { AnyValue, WholeNumber, Decimal, Date, Time, TextLength, List, Custom }
    public enum XLOperator { EqualTo, NotEqualTo, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan, Between, NotBetween }
    public interface IXLDataValidation
    {
        IXLRanges Ranges { get; set; }
        //void Delete();
        //void CopyFrom(IXLDataValidation dataValidation);
        Boolean ShowInputMessage { get; set; }
        Boolean ShowErrorMessage { get; set; }
        Boolean IgnoreBlanks { get; set; }
        Boolean InCellDropdown { get; set; }
        String InputTitle { get; set; }
        String InputMessage { get; set; }
        String ErrorTitle { get; set; }
        String ErrorMessage { get; set; }
        XLErrorStyle ErrorStyle { get; set; }
        XLAllowedValues AllowedValues { get; set; }
        XLOperator Operator { get; set; }

        String Value { get; set; }
        String MinValue { get; set; }
        String MaxValue { get; set; }

        XLWholeNumberCriteria WholeNumber { get; }
        XLDecimalCriteria Decimal { get; }
        XLDateCriteria Date { get; }
        XLTimeCriteria Time { get; }
        XLTextLengthCriteria TextLength { get; }

        void List(String list);
        void List(String list, Boolean inCellDropdown);
        void List(IXLRange range);
        void List(IXLRange range, Boolean inCellDropdown);

        void Custom(String customValidation);
        void Clear();
        Boolean IsDirty();
    }
}
